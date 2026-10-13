<?php

namespace App\Import;

/** Conversions de valeurs de l'ancienne base (Access/Jet) vers les types de la nouvelle application. */
final class Transformer
{
    public static function apply(string $type, mixed $value): mixed
    {
        if (null === $value || (is_string($value) && '' === trim($value))) {
            return null;
        }

        return match ($type) {
            'date', 'date_immutable' => self::date($value),
            'bool', 'boolean' => self::bool($value),
            'int', 'integer' => is_numeric($value) ? (int) $value : null,
            'deliberation' => self::deliberation($value),
            'sexe' => self::sexe($value),
            default => trim((string) $value),
        };
    }

    private static function date(mixed $v): ?\DateTimeImmutable
    {
        if ($v instanceof \DateTimeInterface) {
            return \DateTimeImmutable::createFromInterface($v)->setTime(0, 0);
        }
        $s = trim((string) $v);
        if ('' === $s || str_starts_with($s, '0000') || str_starts_with($s, '1899') || str_starts_with($s, '1900-01-00')) {
            return null;
        }
        try {
            // Formats rencontrés : 2004-06-10 00:00:00 (mdbtools), 10/06/2004 (saisie manuelle)
            $d = preg_match('#^(\d{2})/(\d{2})/(\d{4})#', $s, $m) ? new \DateTimeImmutable("$m[3]-$m[2]-$m[1]") : new \DateTimeImmutable($s);
        } catch (\Exception) {
            return null;
        }

        return $d->setTime(0, 0);
    }

    private static function bool(mixed $v): bool
    {
        return in_array(strtolower(trim((string) $v)), ['1', '-1', 'true', 'vrai', 'oui', 'o', 'y', 'yes', 't', 'x'], true);
    }

    /** 'A' (accord) / 'R' (refus) pour les recours ; entiers 0-3 pour les commissions. */
    private static function deliberation(mixed $v): int
    {
        $s = strtoupper(trim((string) $v));

        return match (true) {
            'A' === $s => 3,
            'R' === $s => 1,
            is_numeric($s) && (int) $s >= 0 && (int) $s <= 3 => (int) $s,
            default => 0,
        };
    }

    /** 'M'/'F'/'I' ou 1-3 → 2 masculin, 3 féminin, 1 indéterminé. */
    private static function sexe(mixed $v): int
    {
        $s = strtoupper(trim((string) $v));

        return match (true) {
            in_array($s, ['M', 'G', '2'], true) => 2,
            in_array($s, ['F', '3'], true) => 3,
            default => 1,
        };
    }
}
