<?php

namespace App\Import;

use App\Entity\Demande;
use App\Entity\Enum\TypeIntervenant;
use App\Entity\Intervenant;
use App\Entity\LegacyMap;
use App\Service\AgrementCalculator;
use Doctrine\DBAL\Connection;
use Doctrine\ORM\EntityManagerInterface;
use Doctrine\ORM\Mapping\ClassMetadata;
use Symfony\Component\PropertyAccess\PropertyAccess;

/** Importe les données de l'ancienne base (copie MySQL de la base Access) dans la base de l'application. */
final class LegacyImporter
{
    private const BATCH = 200;

    public function __construct(
        private readonly EntityManagerInterface $em,
        private readonly AgrementCalculator $agrement,
    ) {
    }

    /**
     * @param list<Step> $steps
     *
     * @return list<StepReport>
     */
    public function import(Connection $source, array $steps, bool $checkOnly = false, bool $strict = false, bool $dryRun = false): array
    {
        $tables = [];
        foreach ($source->createSchemaManager()->listTableNames() as $name) {
            $tables[strtolower($name)] = $name;
        }
        $ctx = new ImportContext($this->em, $source, $tables);
        $reports = [];

        foreach ($steps as $step) {
            $reports[] = $this->plan($step, $ctx);
        }
        if ($strict) {
            foreach ($reports as $r) {
                if ($r->tableAbsente || $r->colonnesAbsentes) {
                    throw new \RuntimeException(sprintf('Étape « %s » : %s', $r->step->name, $r->tableAbsente ? 'table '.$r->step->table.' absente' : 'colonnes absentes : '.implode(', ', $r->colonnesAbsentes)));
                }
            }
        }
        if ($checkOnly) {
            return $reports;
        }

        $this->em->getConnection()->beginTransaction();
        try {
            $this->chargerCorrespondances($ctx);
            foreach ($reports as $report) {
                if (!$report->tableAbsente) {
                    $this->importerEtape($report, $ctx);
                }
            }
            $this->typerIntervenants($ctx);
            $this->recalculerAgrements();
            $dryRun ? $this->em->getConnection()->rollBack() : $this->em->getConnection()->commit();
        } catch (\Throwable $e) {
            $this->em->getConnection()->rollBack();
            throw $e;
        }

        return $reports;
    }

    private function plan(Step $step, ImportContext $ctx): StepReport
    {
        $report = new StepReport($step);
        $real = $ctx->tables[strtolower($step->table)] ?? null;
        if (null === $real) {
            $report->tableAbsente = true;

            return $report;
        }
        $report->table = $real;
        $report->colonnes = array_map('strtolower', array_keys($ctx->source->createSchemaManager()->listTableColumns($real)));

        $voulues = [...$step->key, ...$step->requires, ...array_map(static fn ($f) => is_array($f) ? $f[0] : $f, $step->fields), ...array_map(static fn ($r) => $r[0], $step->refs)];
        foreach (array_unique($voulues) as $col) {
            if (!in_array(strtolower($col), $report->colonnes, true)) {
                $report->colonnesAbsentes[] = $col;
            }
        }
        $report->nbLignes = (int) $ctx->source->fetchOne('SELECT COUNT(*) FROM '.$ctx->source->quoteSingleIdentifier($real));

        return $report;
    }

    private function chargerCorrespondances(ImportContext $ctx): void
    {
        foreach ($this->em->getRepository(LegacyMap::class)->findAll() as $m) {
            $ctx->map[$m->getEntity().'|'.$m->getLegacyKey()] = $m->getNewId();
        }
    }

    private function importerEtape(StepReport $report, ImportContext $ctx): void
    {
        $step = $report->step;
        $meta = $this->em->getClassMetadata($step->entity);
        $accessor = PropertyAccess::createPropertyAccessor();
        $cols = array_flip($report->colonnes);
        $pending = [];

        $flush = function () use (&$pending, $ctx, $step): void {
            $this->em->flush();
            foreach ($pending as [$entity, $key]) {
                $ctx->map[$step->scope().'|'.$key] = $entity->getId();
                $this->em->persist(new LegacyMap($step->scope(), $key, $entity->getId()));
            }
            $this->em->flush();
            $pending = [];
            $this->em->clear();
        };

        $query = 'SELECT * FROM '.$ctx->source->quoteSingleIdentifier($report->table);
        foreach ($ctx->source->iterateAssociative($query) as $raw) {
            $row = array_change_key_case($raw, CASE_LOWER);
            ++$report->lues;

            foreach ($step->requires as $col) {
                if (!isset($cols[strtolower($col)]) || null === $row[strtolower($col)] || '' === $row[strtolower($col)]) {
                    ++$report->ignorees;
                    continue 2;
                }
            }

            $keyCols = array_map('strtolower', $step->key);
            $key = $keyCols && !array_diff($keyCols, array_keys($cols)) && null !== ($row[$keyCols[0]] ?? null)
                ? implode('|', array_map(static fn ($c) => trim((string) $row[$c]), $keyCols))
                : md5(json_encode($row, JSON_THROW_ON_ERROR));

            $extra = $step->map ? ($step->map)($row, $ctx) : [];
            if (null === $extra) {
                ++$report->ignorees;
                continue;
            }

            $id = $ctx->map[$step->scope().'|'.$key] ?? ($extra['id'] ?? null);
            $entity = $id ? $this->em->find($step->entity, $id) : null;
            $nouveau = null === $entity;
            $entity ??= new ($step->entity)();

            foreach ($step->fields as $prop => $spec) {
                [$col, $type] = is_array($spec) ? $spec : [$spec, null];
                $col = strtolower($col);
                if (!isset($cols[$col])) {
                    continue;
                }
                $this->affecter($entity, $meta, $accessor, $prop, $row[$col], $type);
            }
            foreach ($step->defaults as $prop => $value) {
                $accessor->setValue($entity, $prop, $value);
            }
            foreach ($extra as $prop => $value) {
                $this->affecter($entity, $meta, $accessor, $prop, $value, null);
            }

            $orphelin = false;
            foreach ($step->refs as $prop => [$col, $class]) {
                $col = strtolower($col);
                $value = $cols && isset($cols[$col]) ? $row[$col] : null;
                if (null === $value || '' === trim((string) $value)) {
                    $target = null;
                } else {
                    $target = $ctx->resolve($class, trim((string) $value));
                    if (null === $target) {
                        ++$report->refsIntrouvables;
                    }
                }
                $mapping = $meta->getAssociationMapping($prop);
                if (null === $target && !($mapping->joinColumns[0]->nullable ?? true)) {
                    $orphelin = true;
                    break;
                }
                $accessor->setValue($entity, $prop, $target);
            }
            if ($orphelin) {
                ++$report->ignorees;
                continue;
            }

            if ($step->after) {
                ($step->after)($entity, $row, $ctx);
            }

            if ($nouveau) {
                $this->em->persist($entity);
                $pending[] = [$entity, $key];
                ++$report->crees;
            } else {
                ++$report->misAJour;
            }
            if (count($pending) >= self::BATCH) {
                $flush();
            }
        }
        $flush();
    }

    private function affecter(object $entity, ClassMetadata $meta, \Symfony\Component\PropertyAccess\PropertyAccessorInterface $accessor, string $prop, mixed $value, ?string $transformation): void
    {
        $mapping = $meta->hasField($prop) ? $meta->getFieldMapping($prop) : null;
        $type = $transformation ?? ($mapping?->type ?? 'string');
        $value = Transformer::apply($type, $value);

        if (null === $value && null !== $mapping && !$mapping->nullable) {
            $value = match ($mapping->type) {
                'boolean' => false,
                'integer' => null,
                default => 'string' === $mapping->type || 'text' === $mapping->type ? '' : null,
            };
            if (null === $value) {
                return;
            }
        }
        $accessor->setValue($entity, $prop, $value);
    }

    private function typerIntervenants(ImportContext $ctx): void
    {
        foreach ($ctx->typesIntervenants() as $id => $types) {
            $intervenant = $this->em->find(Intervenant::class, $id);
            if (null === $intervenant) {
                continue;
            }
            $psychologue = ($types['psy'] ?? false) && !($types['as'] ?? false);
            $intervenant->setType(($psychologue ? TypeIntervenant::Psychologue : TypeIntervenant::AssistanteSociale)->value);
        }
        $this->em->flush();
        $this->em->clear();
    }

    /** Seules les demandes ayant des décisions sont recalculées : les autres conservent l'agrément importé. */
    private function recalculerAgrements(): void
    {
        $ids = $this->em->createQuery('SELECT DISTINCT IDENTITY(d.demande) FROM App\Entity\Decision d')->getSingleColumnResult();
        foreach (array_chunk($ids, self::BATCH) as $chunk) {
            foreach ($this->em->getRepository(Demande::class)->findBy(['id' => $chunk]) as $demande) {
                $this->agrement->actualiser($demande);
            }
            $this->em->flush();
            $this->em->clear();
        }
    }
}
