<?php

namespace App\Repository;

use App\Entity\Dossier;
use Doctrine\Bundle\DoctrineBundle\Repository\ServiceEntityRepository;
use Doctrine\Persistence\ManagerRegistry;

/** @extends ServiceEntityRepository<Dossier> */
class DossierRepository extends ServiceEntityRepository
{
    public function __construct(ManagerRegistry $registry)
    {
        parent::__construct($registry, Dossier::class);
    }

    /** @return list<Dossier> */
    public function rechercher(string $q): array
    {
        $qb = $this->createQueryBuilder('d')->orderBy('d.numero');
        if ('' !== $q) {
            $qb->leftJoin('App\Entity\Individu', 'i', 'WITH', 'i.dossier = d')
                ->andWhere('d.numero LIKE :q OR i.nom LIKE :q OR i.prenom LIKE :q OR d.ville LIKE :q')
                ->setParameter('q', '%'.$q.'%')
                ->distinct();
        }

        return $qb->getQuery()->getResult();
    }
}
