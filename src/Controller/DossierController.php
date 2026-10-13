<?php

namespace App\Controller;

use App\Entity\Demande;
use App\Entity\Dossier;
use App\Form\EntityType;
use App\Repository\DossierRepository;
use App\Service\AgrementCalculator;
use App\Service\EcheancierCalculator;
use Doctrine\ORM\EntityManagerInterface;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Attribute\Route;

class DossierController extends AbstractController
{
    public function __construct(private readonly EntityManagerInterface $em)
    {
    }

    #[Route('/', name: 'home')]
    public function home(): Response
    {
        $repo = $this->em->getRepository(Demande::class);

        return $this->render('home.html.twig', [
            'nbDossiers' => $this->em->getRepository(Dossier::class)->count([]),
            'nbDemandes' => $repo->count([]),
            'aActualiser' => $repo->findBy(['aActualiser' => true, 'dateSortie' => null], ['dateDemande' => 'ASC']),
        ]);
    }

    #[Route('/dossiers', name: 'dossier_index')]
    public function index(Request $request, DossierRepository $repository): Response
    {
        $q = trim((string) $request->query->get('q', ''));

        return $this->render('dossier/index.html.twig', ['dossiers' => $repository->rechercher($q), 'q' => $q]);
    }

    #[Route('/dossiers/nouveau', name: 'dossier_new', methods: ['GET', 'POST'])]
    public function new(Request $request): Response
    {
        return $this->form($request, new Dossier());
    }

    #[Route('/dossiers/{id}', name: 'dossier_show', requirements: ['id' => '\d+'])]
    public function show(Dossier $dossier): Response
    {
        return $this->render('dossier/show.html.twig', ['dossier' => $dossier]);
    }

    #[Route('/dossiers/{id}/modifier', name: 'dossier_edit', requirements: ['id' => '\d+'], methods: ['GET', 'POST'])]
    public function edit(Request $request, Dossier $dossier): Response
    {
        return $this->form($request, $dossier);
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
        return $this->render('demande/show.html.twig', ['demande' => $demande, 'echeancier' => $echeancier->calculer($demande)]);
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

    private function form(Request $request, Dossier $dossier): Response
    {
        $form = $this->createForm(EntityType::class, $dossier, ['data_class' => Dossier::class]);
        $form->handleRequest($request);
        if ($form->isSubmitted() && $form->isValid()) {
            $this->em->persist($dossier);
            $this->em->flush();
            $this->addFlash('success', 'Dossier enregistré.');

            return $this->redirectToRoute('dossier_show', ['id' => $dossier->getId()]);
        }

        return $this->render('dossier/form.html.twig', ['form' => $form, 'dossier' => $dossier], new Response(status: $form->isSubmitted() ? 422 : 200));
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
}
