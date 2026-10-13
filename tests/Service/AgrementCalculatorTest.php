<?php

namespace App\Tests\Service;

use App\Entity\Decision;
use App\Entity\Enum\Deliberation;
use App\Entity\TypeRecours;
use App\Service\AgrementCalculator;
use PHPUnit\Framework\TestCase;

class AgrementCalculatorTest extends TestCase
{
    private function d(string $date): \DateTimeImmutable
    {
        return new \DateTimeImmutable($date);
    }

    private function commission(string $date, Deliberation $d, ?string $signature = null): Decision
    {
        return (new Decision())->setProvenance(-2)->setCommissionInitiale($this->d($date))->setDateCommission($this->d($date))
            ->setDateSignature($signature ? $this->d($signature) : null)->setDeliberation($d->value);
    }

    private function recoursContentieux(string $commissionInitiale, Deliberation $d, ?string $agrement, bool $conseilEtat): Decision
    {
        return (new Decision())->setProvenance(5)->setTypeRecours((new TypeRecours())->setConseilEtat($conseilEtat))
            ->setCommissionInitiale($this->d($commissionInitiale))->setDateSignature($agrement ? $this->d($agrement) : null)
            ->setDateDepot($this->d($commissionInitiale.' +1 month'))->setDeliberation($d->value);
    }

    public function testAucuneDecision(): void
    {
        $r = (new AgrementCalculator())->calculer([]);
        self::assertNull($r->dateAgrement);
        self::assertNull($r->dateRetrait);
    }

    public function testAccordDeLaCommissionInitiale(): void
    {
        $r = (new AgrementCalculator())->calculer([$this->commission('2020-03-10', Deliberation::Accord)]);
        self::assertEquals($this->d('2020-03-10'), $r->dateAgrement);
        self::assertSame(-2, $r->provenance);
    }

    public function testRefusPuisAccordSurRecoursContentieux(): void
    {
        $r = (new AgrementCalculator())->calculer([
            $this->commission('2020-03-10', Deliberation::Refus),
            $this->recoursContentieux('2020-03-10', Deliberation::Accord, '2021-01-15', false),
        ]);
        self::assertEquals($this->d('2021-01-15'), $r->dateAgrement);
        self::assertSame(5, $r->provenance);
    }

    public function testRefusConfirmeParLeConseilEtat(): void
    {
        $r = (new AgrementCalculator())->calculer([
            $this->commission('2020-03-10', Deliberation::Accord),
            $this->recoursContentieux('2020-03-10', Deliberation::Refus, null, true),
        ]);
        self::assertNull($r->dateAgrement);
    }

    public function testRetraitPuisRecoursFavorableSuspendLeRetrait(): void
    {
        $retrait = $this->commission('2022-05-01', Deliberation::Retrait);
        $premiere = $this->commission('2020-03-10', Deliberation::Accord);
        $recours = $this->recoursContentieux('2022-05-01', Deliberation::Accord, '2023-02-01', false);

        $seul = (new AgrementCalculator())->calculer([$premiere, $retrait]);
        self::assertEquals($this->d('2020-03-10'), $seul->dateAgrement);
        self::assertEquals($this->d('2022-05-01'), $seul->dateRetrait);

        $apres = (new AgrementCalculator())->calculer([$premiere, $retrait, $recours]);
        self::assertNull($apres->dateRetrait);
    }

    public function testCommissionSansDateUtiliseLaSignature(): void
    {
        $recours = (new Decision())->setProvenance(-1)->setCommissionInitiale($this->d('2020-03-10'))->setDateSignature($this->d('2020-06-01'))
            ->setDateDepot($this->d('2020-04-01'))->setDeliberation(Deliberation::Accord->value);
        $r = (new AgrementCalculator())->calculer([$this->commission('2020-03-10', Deliberation::Refus), $recours]);
        self::assertEquals($this->d('2020-06-01'), $r->dateAgrement);
        self::assertSame(-1, $r->provenance);
    }
}
