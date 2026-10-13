<?php

namespace App\Service;

final class AgrementResult
{
    public function __construct(
        public readonly ?\DateTimeImmutable $dateAgrement,
        public readonly ?int $provenance,
        public readonly ?\DateTimeImmutable $dateRetrait,
    ) {
    }
}
