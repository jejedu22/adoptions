<?php

namespace App\Controller;

use App\Entity\Demande;
use App\Entity\Dossier;
use App\Entity\HistoriqueIntervenant;
use App\Entity\Individu;
use App\Form\EntityType;
use App\Form\NouveauDossier;
use App\Form\NouveauDossierType;
use App\Repository\DossierRepository;
use App\Service\AgrementCalculator;
use App\Service\EcheancierCalculator;
use App\Service\Statut;
use App\Service\StatutCalculator;
use App\Service\TacheFinder;
use Doctrine\ORM\EntityManagerInterface;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Attribute\Route;

class DossierController extends AbstractController
{
    private const FILTRES = [
        'tous' => 'Tous',
        'instruction' => 'En instruction',
        'agree' => 'Agréés',
        'refus' => 'Refus / retrait',
        'sorti' => 'Sortis',
        'atraiter' => '⚠ À traiter',
    ];

    public function __construct(
        private readonly EntityManagerInterface $em,
        private readonly StatutCalculator $statuts,
        private readonly TacheFinder $taches,
    ) {
    }

    #[Route('/', name: 'home')]
    public function home(Request $request): Response
    {
        $aTraiter = $this->taches->trouver(30);
        $retards = array_filter($aTraiter, static fn ($t) => $t->enRetard());
        $actifs = 0;
        foreach ($this->em->getRepository(Dossier::class)->findAll() as $dossier) {
            $d = StatutCalculator::demandeCourante($dossier->getDemandes());
            if (null !== $d && null === $d->getDateSortie()) {
                ++$actifs;
            }
        }

        $recents = [];
        $ids = $request->getSession()->get('dossiers_recents', []);
        if ($ids) {
            $trouves = $this->em->getRepository(Dossier::class)->findBy(['id' => $ids]);
            usort($trouves, static fn ($a, $b) => array_search($a->getId(), $ids) <=> array_search($b->getId(), $ids));
            $recents = $trouves;
        }

        return $this->render('home.html.twig', [
            'actifs' => $actifs,
            'retards' => count($retards),
            'reactualisations' => count(array_filter($aTraiter, static fn ($t) => 'reactualisation' === $t->type)),
            'commissions' => count(array_filter($aTraiter, static fn ($t) => 'decision' === $t->type)),
            'taches' => array_slice($aTraiter, 0, 5),
            'totalTaches' => count($aTraiter),
            'recents' => $recents,
        ]);
    }

    #[Route('/dossiers', name: 'dossier_index')]
    public function index(Request $request, DossierRepository $repository): Response
    {
        $q = trim((string) $request->query->get('q', ''));
        $filtre = (string) $request->query->get('statut', 'tous');
        $filtre = isset(self::FILTRES[$filtre]) ? $filtre : 'tous';

        $enRetard = [];
        foreach ($this->taches->trouver(0) as $t) {
            if ($t->enRetard()) {
                $enRetard[$t->demande->getDossier()->getId()] = true;
            }
        }

        $lignes = [];
        $compteurs = array_fill_keys(array_keys(self::FILTRES), 0);
        foreach ($repository->rechercher($q) as $dossier) {
            $demande = StatutCalculator::demandeCourante($dossier->getDemandes());
            $statut = $demande ? $this->statuts->calculer($demande) : null;
            $categories = ['tous'];
            if (null !== $statut) {
                $categories[] = match ($statut->libelle) {
                    'En instruction', 'Suspendu' => 'instruction',
                    'Agréé', 'À réactualiser' => 'agree',
                    'Refus', 'Refus · recours', 'Retrait' => 'refus',
                    'Sorti' => 'sorti',
                    default => 'tous',
                };
            }
            if (isset($enRetard[$dossier->getId()])) {
                $categories[] = 'atraiter';
            }
            foreach ($categories as $c) {
                ++$compteurs[$c];
            }
            if (in_array($filtre, $categories, true)) {
                $lignes[] = ['dossier' => $dossier, 'demande' => $demande, 'statut' => $statut, 'retard' => isset($enRetard[$dossier->getId()])];
            }
        }

        return $this->render('dossier/index.html.twig', [
            'lignes' => $lignes, 'q' => $q, 'filtre' => $filtre, 'filtres' => self::FILTRES, 'compteurs' => $compteurs,
        ]);
    }

    #[Route('/dossiers/nouveau', name: 'dossier_new', methods: ['GET', 'POST'])]
    public function new(Request $request): Response
    {
        $data = new NouveauDossier();
        $data->dateDemande = new \DateTimeImmutable('today');
        $data->numero = $this->numeroSuivant();
        $form = $this->createForm(NouveauDossierType::class, $data);
        $form->handleRequest($request);

        if ($form->isSubmitted() && $form->isValid()) {
            $dossier = $this->creerDossier($data);

            $this->addFlash('success', 'Le dossier '.$dossier->getNumero().' a été créé.');

            return $this->redirectToRoute('dossier_show', ['id' => $dossier->getId()]);
        }

        return $this->render('dossier/wizard.html.twig', ['form' => $form], new Response(status: $form->isSubmitted() ? 422 : 200));
    }

    #[Route('/dossiers/{id}', name: 'dossier_show', requirements: ['id' => '\d+'])]
    public function show(Request $request, Dossier $dossier, EcheancierCalculator $echeancier): Response
    {
        $session = $request->getSession();
        $recents = array_values(array_unique([$dossier->getId(), ...$session->get('dossiers_recents', [])]));
        $session->set('dossiers_recents', array_slice($recents, 0, 5));

        $demande = StatutCalculator::demandeCourante($dossier->getDemandes());
        $statut = $demande ? $this->statuts->calculer($demande) : null;
        $demandeur = null;
        foreach ($dossier->getIndividus() as $i) {
            $demandeur ??= $i;
            if ('DE' === $i->getType()) {
                $demandeur = $i;
                break;
            }
        }

        $tachesDossier = array_values(array_filter($this->taches->trouver(90), static fn ($t) => $t->demande->getDossier() === $dossier));
        $placements = [];
        foreach ($dossier->getDemandes() as $d) {
            foreach ($d->getPlacements() as $p) {
                $placements[] = $p;
            }
        }

        return $this->render('dossier/show.html.twig', [
            'dossier' => $dossier, 'demande' => $demande, 'statut' => $statut, 'demandeur' => $demandeur,
            'echeancier' => $demande ? $echeancier->calculer($demande) : [], 'taches' => $tachesDossier,
            'placements' => $placements,
            'statuts' => array_map(fn (Demande $d) => $this->statuts->calculer($d), $dossier->getDemandes()->toArray()),
        ]);
    }

    #[Route('/dossiers/{id}/modifier', name: 'dossier_edit', requirements: ['id' => '\d+'], methods: ['GET', 'POST'])]
    public function edit(Request $request, Dossier $dossier): Response
    {
        $form = $this->createForm(EntityType::class, $dossier, ['data_class' => Dossier::class]);
        $form->handleRequest($request);
        if ($form->isSubmitted() && $form->isValid()) {
            $this->em->flush();
            $this->addFlash('success', 'Dossier enregistré.');

            return $this->redirectToRoute('dossier_show', ['id' => $dossier->getId()]);
        }

        return $this->render('dossier/form.html.twig', ['form' => $form, 'dossier' => $dossier], new Response(status: $form->isSubmitted() ? 422 : 200));
    }

    #[Route('/dossiers/{id}/supprimer', name: 'dossier_delete', requirements: ['id' => '\d+'], methods: ['POST'])]
    public function delete(Request $request, Dossier $dossier): Response
    {
        if ($this->isCsrfTokenValid('delete_dossier'.$dossier->getId(), (string) $request->request->get('_token'))) {
            $this->em->remove($dossier);
            $this->em->flush();
            $this->addFlash('success', 'Dossier supprimé.');
        }

        return $this->redirectToRoute('dossier_index');
    }

    #[Route('/demandes/nouvelle/{dossier}', name: 'demande_new', requirements: ['dossier' => '\d+'], methods: ['GET', 'POST'])]
    public function newDemande(Request $request, Dossier $dossier): Response
    {
        $demande = (new Demande())->setDossier($dossier)->setDateDemande(new \DateTimeImmutable('today'));
        $numero = 1;
        foreach ($dossier->getDemandes() as $d) {
            $numero = max($numero, ($d->getNumero() ?? 0) + 1);
        }
        $demande->setNumero($numero);

        return $this->formDemande($request, $demande);
    }

    #[Route('/demandes/{id}', name: 'demande_show', requirements: ['id' => '\d+'])]
    public function showDemande(Demande $demande, EcheancierCalculator $echeancier): Response
    {
        return $this->render('demande/show.html.twig', [
            'demande' => $demande, 'echeancier' => $echeancier->calculer($demande), 'statut' => $this->statuts->calculer($demande),
        ]);
    }

    #[Route('/demandes/{id}/modifier', name: 'demande_edit', requirements: ['id' => '\d+'], methods: ['GET', 'POST'])]
    public function editDemande(Request $request, Demande $demande): Response
    {
        return $this->formDemande($request, $demande);
    }

    #[Route('/demandes/{id}/recalculer', name: 'demande_recalcul', requirements: ['id' => '\d+'], methods: ['POST'])]
    public function recalculer(Request $request, Demande $demande, AgrementCalculator $agrement): Response
    {
        if ($this->isCsrfTokenValid('recalcul'.$demande->getId(), (string) $request->request->get('_token'))) {
            $agrement->actualiser($demande);
            $this->em->flush();
            $this->addFlash('success', 'Agrément recalculé.');
        }

        return $this->redirectToRoute('demande_show', ['id' => $demande->getId()]);
    }

    /** Action rapide : enregistre la réception d'un rapport à la date du jour. */
    #[Route('/interventions/{id}/recu', name: 'intervention_recu', requirements: ['id' => '\d+'], methods: ['POST'])]
    public function recu(Request $request, HistoriqueIntervenant $intervention): Response
    {
        if ($this->isCsrfTokenValid('recu'.$intervention->getId(), (string) $request->request->get('_token'))) {
            $intervention->setDateRecu(new \DateTimeImmutable('today'));
            $this->em->flush();
            $this->addFlash('success', 'Réception du rapport enregistrée.');
        }
        $retour = (string) $request->request->get('retour', '');

        return str_starts_with($retour, '/') && !str_starts_with($retour, '//')
            ? $this->redirect($retour)
            : $this->redirectToRoute('demande_show', ['id' => $intervention->getDemande()->getId()]);
    }

    private function formDemande(Request $request, Demande $demande): Response
    {
        $form = $this->createForm(EntityType::class, $demande, ['data_class' => Demande::class, 'exclude' => 'dossier']);
        $form->handleRequest($request);
        if ($form->isSubmitted() && $form->isValid()) {
            $this->em->persist($demande);
            $this->em->flush();
            $this->addFlash('success', 'Demande enregistrée.');

            return $this->redirectToRoute('demande_show', ['id' => $demande->getId()]);
        }

        return $this->render('demande/form.html.twig', ['form' => $form, 'demande' => $demande], new Response(status: $form->isSubmitted() ? 422 : 200));
    }

    private function creerDossier(NouveauDossier $n): Dossier
    {
        $dossier = (new Dossier())
            ->setNumero((string) $n->numero)
            ->setAdresse1($n->adresse1)->setCodePostal($n->codePostal)->setVille($n->ville)->setTelephone($n->telephone)
            ->setCommune($n->commune)->setCanton($n->commune?->getCanton())->setCirconscription($n->commune?->getCirconscription())
            ->setSituationFamiliale($n->situationFamiliale)->setDateMariage($n->dateMariage)->setLieuMariage($n->lieuMariage);
        if (null === $dossier->getCodePostal() && $n->commune) {
            $dossier->setCodePostal($n->commune->getCodePostal());
        }
        if (null === $dossier->getVille() && $n->commune) {
            $dossier->setVille($n->commune->getNom());
        }

        $demandeur = (new Individu())->setDossier($dossier)->setType('DE')->setNom((string) $n->nom)->setPrenom($n->prenom)
            ->setDateNaissance($n->dateNaissance)->setProfession($n->profession);
        $dossier->addIndividu($demandeur);
        if ($n->conjoint && $n->conjointNom) {
            $dossier->addIndividu((new Individu())->setDossier($dossier)->setType('CO')->setNom($n->conjointNom)->setPrenom($n->conjointPrenom)->setDateNaissance($n->conjointDateNaissance));
        }

        $demande = (new Demande())->setDossier($dossier)->setNumero(1)->setDateDemande($n->dateDemande)
            ->setNature($n->nature)->setType($n->type)->setNombreEnfants($n->nombreEnfants)->setOrigineEthnique($n->origineEthnique)->setFratrie($n->fratrie);
        $dossier->addDemande($demande);

        $this->em->persist($dossier);
        $this->em->flush();

        return $dossier;
    }

    private function numeroSuivant(): string
    {
        $max = 0;
        foreach ($this->em->createQuery('SELECT d.numero FROM App\Entity\Dossier d')->getSingleColumnResult() as $numero) {
            if (ctype_digit((string) $numero)) {
                $max = max($max, (int) $numero);
            }
        }

        return $max > 0 ? str_pad((string) ($max + 1), strlen((string) $max), '0', STR_PAD_LEFT) : date('Y').'001';
    }
}
