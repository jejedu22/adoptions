<?php

namespace App\Controller;

use App\Entity\Decision;
use App\Entity\Placement;
use App\Service\Catalogue;
use App\Service\TacheFinder;
use Doctrine\ORM\EntityManagerInterface;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Attribute\Route;

/** Pages transverses : échéances à traiter, commissions, placements, recherche et administration. */
class TableauController extends AbstractController
{
    public function __construct(private readonly EntityManagerInterface $em, private readonly TacheFinder $taches)
    {
    }

    #[Route('/a-traiter', name: 'taches')]
    public function taches(Request $request): Response
    {
        $filtre = (string) $request->query->get('filtre', 'retard');
        $toutes = $this->taches->trouver(30);
        $filtres = [
            'retard' => ['En retard', static fn ($t) => $t->enRetard()],
            'semaine' => ['7 prochains jours', static fn ($t) => !$t->enRetard() && $t->retard >= -7],
            'mois' => ['30 jours', static fn ($t) => !$t->enRetard()],
            'tout' => ['Tout', static fn ($t) => true],
        ];
        $filtre = isset($filtres[$filtre]) ? $filtre : 'retard';
        $compteurs = array_map(static fn ($f) => count(array_filter($toutes, $f[1])), $filtres);

        return $this->render('taches.html.twig', [
            'taches' => array_values(array_filter($toutes, $filtres[$filtre][1])),
            'filtre' => $filtre,
            'filtres' => array_map(static fn ($f) => $f[0], $filtres),
            'compteurs' => $compteurs,
        ]);
    }

    #[Route('/commissions', name: 'commissions')]
    public function commissions(): Response
    {
        $decisions = $this->em->createQuery(
            'SELECT x, d FROM App\Entity\Decision x JOIN x.demande d ORDER BY x.dateCommission DESC, x.id DESC'
        )->setMaxResults(300)->getResult();
        $aVenir = array_filter($decisions, static fn (Decision $x) => 0 === $x->getDeliberation());
        $passees = array_filter($decisions, static fn (Decision $x) => 0 !== $x->getDeliberation());

        return $this->render('commissions.html.twig', ['aVenir' => $aVenir, 'passees' => array_slice($passees, 0, 50)]);
    }

    #[Route('/placements', name: 'placements')]
    public function placements(): Response
    {
        $placements = $this->em->getRepository(Placement::class)->findBy([], ['datePlacement' => 'DESC'], 300);

        return $this->render('placements.html.twig', ['placements' => $placements]);
    }

    #[Route('/recherche', name: 'recherche')]
    public function recherche(Request $request): Response
    {
        $q = trim((string) $request->query->get('q', ''));
        $dossiers = $demandes = [];
        if (mb_strlen($q) >= 2) {
            $like = '%'.$q.'%';
            $dossiers = $this->em->createQuery(
                'SELECT DISTINCT d FROM App\Entity\Dossier d LEFT JOIN d.individus i
                 WHERE d.numero LIKE :q OR i.nom LIKE :q OR i.prenom LIKE :q OR d.ville LIKE :q OR d.telephone LIKE :q ORDER BY d.numero'
            )->setParameter('q', $like)->setMaxResults(50)->getResult();
            if (ctype_digit($q)) {
                $demandes = $this->em->createQuery('SELECT d FROM App\Entity\Demande d WHERE d.numero = :n')->setParameter('n', (int) $q)->setMaxResults(20)->getResult();
            }
        }

        return $this->render('recherche.html.twig', ['q' => $q, 'dossiers' => $dossiers, 'demandes' => $demandes]);
    }

    #[Route('/administration', name: 'administration')]
    public function administration(): Response
    {
        return $this->render('administration.html.twig', ['menus' => Catalogue::menus()]);
    }
}
