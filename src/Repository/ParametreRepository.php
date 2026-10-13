<?php

namespace App\Repository;

use App\Entity\Parametre;
use Doctrine\Bundle\DoctrineBundle\Repository\ServiceEntityRepository;
use Doctrine\Persistence\ManagerRegistry;

/** @extends ServiceEntityRepository<Parametre> */
class ParametreRepository extends ServiceEntityRepository
{
    public function __construct(ManagerRegistry $registry)
    {
        parent::__construct($registry, Parametre::class);
    }

    /** @return array<int, int> */
    public function moisParId(): array
    {
        $r = [];
        foreach ($this->findAll() as $p) {
            $r[$p->getId()] = $p->getNombreMois();
        }

        return $r;
    }
}
