<?php

namespace App\Controller;

use App\Entity\Decision;
use App\Entity\Enum\Deliberation;
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
    public function commissions(Request $request): Response
    {
        $f = [];
        foreach (['dossier', 'origine', 'decision', 'commission_du', 'commission_au', 'prevue_du', 'prevue_au'] as $k) {
            $f[$k] = trim((string) $request->query->get($k, ''));
        }
        $date = static fn (string $v): ?\DateTimeImmutable => preg_match('/^\d{4}-\d{2}-\d{2}$/', $v) ? new \DateTimeImmutable($v) : null;

        $qb = $this->em->createQueryBuilder()->select('x', 'd', 'dos')->from(Decision::class, 'x')
            ->join('x.demande', 'd')->join('d.dossier', 'dos')
            ->orderBy('x.dateCommission', 'DESC')->addOrderBy('x.id', 'DESC');

        if ('' !== $f['dossier']) {
            $qb->andWhere('dos.numero LIKE :dossier OR dos.ville LIKE :dossier OR EXISTS (SELECT 1 FROM App\Entity\Individu i WHERE i.dossier = dos AND (i.nom LIKE :dossier OR i.prenom LIKE :dossier))')
                ->setParameter('dossier', '%'.$f['dossier'].'%');
        }
        match ($f['origine']) {
            'initiale' => $qb->andWhere('x.provenance = -2'),
            'administratif' => $qb->andWhere('x.provenance = -1'),
            'contentieux' => $qb->andWhere('x.provenance > 0'),
            default => null,
        };
        if (ctype_digit($f['decision']) && null !== Deliberation::tryFrom((int) $f['decision'])) {
            $qb->andWhere('x.deliberation = :decision')->setParameter('decision', (int) $f['decision']);
        }
        foreach (['commission' => 'x.dateCommission', 'prevue' => 'x.dateDecisionPrevue'] as $key => $champ) {
            if ($du = $date($f[$key.'_du'])) {
                $qb->andWhere("$champ >= :{$key}_du")->setParameter($key.'_du', $du);
            }
            if ($au = $date($f[$key.'_au'])) {
                $qb->andWhere("$champ <= :{$key}_au")->setParameter($key.'_au', $au);
            }
        }

        $limite = 500;
        $decisions = $qb->setMaxResults($limite + 1)->getQuery()->getResult();

        return $this->render('commissions.html.twig', [
            'decisions' => array_slice($decisions, 0, $limite),
            'tronque' => count($decisions) > $limite,
            'limite' => $limite,
            'f' => $f,
            'filtresActifs' => '' !== implode('', $f),
            'deliberations' => Deliberation::cases(),
        ]);
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
