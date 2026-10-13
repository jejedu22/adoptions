<?php

namespace App\Tests\Controller;

use App\Entity\Dossier;
use Doctrine\ORM\EntityManagerInterface;
use Doctrine\ORM\Tools\SchemaTool;
use Symfony\Bundle\FrameworkBundle\Test\WebTestCase;
use Symfony\Component\Security\Core\User\InMemoryUser;

class SmokeTest extends WebTestCase
{
    public function testParcoursPrincipal(): void
    {
        $client = static::createClient();
        $em = static::getContainer()->get(EntityManagerInterface::class);
        $tool = new SchemaTool($em);
        $tool->dropSchema($em->getMetadataFactory()->getAllMetadata());
        $tool->createSchema($em->getMetadataFactory()->getAllMetadata());

        $client->request('GET', '/');
        self::assertResponseStatusCodeSame(401);

        $client->loginUser(new InMemoryUser('admin', $_SERVER['APP_ADMIN_PASSWORD_HASH'], ['ROLE_USER']));
        $client->request('GET', '/');
        self::assertResponseIsSuccessful();

        foreach (['/dossiers', '/dossiers/nouveau', '/gestion/professions', '/gestion/professions/nouveau', '/gestion/parametres'] as $url) {
            $client->request('GET', $url);
            self::assertResponseIsSuccessful($url);
        }

        $crawler = $client->request('GET', '/dossiers/nouveau');
        $form = $crawler->selectButton('Valider')->form(['entity[numero]' => 'D-001', 'entity[ville]' => 'Nancy']);
        $client->submit($form);
        self::assertResponseRedirects();
        $dossier = $em->getRepository(Dossier::class)->findOneBy(['numero' => 'D-001']);
        self::assertNotNull($dossier);

        $client->request('GET', '/dossiers/'.$dossier->getId());
        self::assertResponseIsSuccessful();

        $crawler = $client->request('GET', '/demandes/nouvelle/'.$dossier->getId());
        $client->submit($crawler->selectButton('Valider')->form());
        self::assertResponseRedirects();
        $client->followRedirect();
        self::assertResponseIsSuccessful();
        self::assertSelectorTextContains('h1', 'Demande n° 1');

        $demande = $dossier->getDemandes()->first() ?: $em->getRepository(\App\Entity\Demande::class)->findOneBy([]);
        $crawler = $client->request('GET', '/gestion/decisions/nouveau?parent='.$demande->getId());
        self::assertResponseIsSuccessful();
        $form = $crawler->selectButton('Valider')->form([
            'entity[dateCommission]' => '2024-03-10',
            'entity[commissionInitiale]' => '2024-03-10',
            'entity[deliberation]' => '3',
            'entity[provenance]' => '-2',
        ]);
        $client->submit($form);
        self::assertResponseRedirects();
        $em->clear();
        $demande = $em->getRepository(\App\Entity\Demande::class)->findOneBy([]);
        self::assertSame('2024-03-10', $demande->getDateAgrement()?->format('Y-m-d'));
    }
}
