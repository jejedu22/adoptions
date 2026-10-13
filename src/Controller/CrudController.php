<?php

namespace App\Controller;

use App\Entity\Decision;
use App\Entity\Demande;
use App\Entity\HistoriqueIntervenant;
use App\Entity\Parametre;
use App\Form\EntityType;
use App\Service\AgrementCalculator;
use App\Service\Catalogue;
use Doctrine\ORM\EntityManagerInterface;
use Doctrine\ORM\Exception\ORMException;
use Doctrine\DBAL\Exception\ForeignKeyConstraintViolationException;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\PropertyAccess\PropertyAccess;
use Symfony\Component\Routing\Attribute\Route;

#[Route('/gestion/{slug}', requirements: ['slug' => '[a-z-]+'])]
class CrudController extends AbstractController
{
    public function __construct(
        private readonly EntityManagerInterface $em,
        private readonly AgrementCalculator $agrement,
    ) {
    }

    #[Route('', name: 'crud_index', methods: ['GET'])]
    public function index(string $slug): Response
    {
        $config = $this->config($slug);
        $items = $this->em->getRepository($config['class'])->findBy([], ['id' => 'ASC']);

        return $this->render('crud/index.html.twig', ['slug' => $slug, 'config' => $config, 'items' => $items]);
    }

    #[Route('/nouveau', name: 'crud_new', methods: ['GET', 'POST'])]
    public function new(Request $request, string $slug): Response
    {
        $config = $this->config($slug);
        $entity = new $config['class']();
        $parent = $this->resolveParent($request, $config, $entity);

        return $this->handle($request, $slug, $config, $entity, $parent);
    }

    #[Route('/{id}/modifier', name: 'crud_edit', requirements: ['id' => '\d+'], methods: ['GET', 'POST'])]
    public function edit(Request $request, string $slug, int $id): Response
    {
        $config = $this->config($slug);
        $entity = $this->em->find($config['class'], $id) ?? throw $this->createNotFoundException();
        $parent = isset($config['parent']) ? PropertyAccess::createPropertyAccessor()->getValue($entity, $config['parent']) : null;

        return $this->handle($request, $slug, $config, $entity, $parent);
    }

    #[Route('/{id}/supprimer', name: 'crud_delete', requirements: ['id' => '\d+'], methods: ['POST'])]
    public function delete(Request $request, string $slug, int $id): Response
    {
        $config = $this->config($slug);
        $entity = $this->em->find($config['class'], $id) ?? throw $this->createNotFoundException();
        if (!$this->isCsrfTokenValid('delete'.$slug.$id, (string) $request->request->get('_token'))) {
            throw $this->createAccessDeniedException();
        }
        $parent = isset($config['parent']) ? PropertyAccess::createPropertyAccessor()->getValue($entity, $config['parent']) : null;

        try {
            $this->em->remove($entity);
            $this->em->flush();
            if ($parent instanceof Demande) {
                $parent->getDecisions()->removeElement($entity);
                $this->agrement->actualiser($parent);
                $this->em->flush();
            }
            $this->addFlash('success', 'Enregistrement supprimé.');
        } catch (ForeignKeyConstraintViolationException) {
            $this->addFlash('danger', 'Suppression refusée : des enregistrements connexes existent.');
        }

        return $this->back($slug, $config, $parent);
    }

    private function handle(Request $request, string $slug, array $config, object $entity, ?object $parent): Response
    {
        $form = $this->createForm(EntityType::class, $entity, [
            'data_class' => $config['class'],
            'exclude' => $config['parent'] ?? null,
            'id_editable' => $entity instanceof Parametre,
        ]);
        $form->handleRequest($request);

        if ($form->isSubmitted() && $form->isValid()) {
            if ($entity instanceof Decision) {
                $this->normaliserProvenance($entity);
            }
            $this->em->persist($entity);
            if ($entity instanceof Decision || $entity instanceof HistoriqueIntervenant) {
                $demande = $entity->getDemande();
                if ($entity instanceof Decision && !$demande->getDecisions()->contains($entity)) {
                    $demande->getDecisions()->add($entity);
                }
                $this->em->flush();
                $this->agrement->actualiser($demande);
            }
            $this->em->flush();
            $this->addFlash('success', 'Enregistrement validé.');

            return $this->back($slug, $config, $parent);
        }

        return $this->render('crud/form.html.twig', [
            'slug' => $slug, 'config' => $config, 'form' => $form, 'entity' => $entity, 'parent' => $parent,
        ], new Response(status: $form->isSubmitted() ? 422 : 200));
    }

    private function normaliserProvenance(Decision $decision): void
    {
        $recours = $decision->getTypeRecours();
        if (null !== $recours) {
            $decision->setProvenance((int) $recours->getId());
        } elseif ($decision->getProvenance() > 0) {
            $decision->setProvenance(AgrementCalculator::PROVENANCE_DEMANDE_INITIALE);
        }
    }

    private function resolveParent(Request $request, array $config, object $entity): ?object
    {
        if (!isset($config['parent'])) {
            return null;
        }
        $parentClass = (new \ReflectionProperty($config['class'], $config['parent']))->getType()->getName();
        $parent = $this->em->find($parentClass, $request->query->getInt('parent')) ?? throw $this->createNotFoundException('Parent manquant.');
        PropertyAccess::createPropertyAccessor()->setValue($entity, $config['parent'], $parent);

        return $parent;
    }

    private function back(string $slug, array $config, ?object $parent): Response
    {
        if (null === $parent) {
            return $this->redirectToRoute('crud_index', ['slug' => $slug]);
        }
        if ($parent instanceof \App\Entity\Placement) {
            return $this->redirectToRoute('demande_show', ['id' => $parent->getDemande()->getId()]);
        }
        if ($parent instanceof Demande) {
            return $this->redirectToRoute('demande_show', ['id' => $parent->getId()]);
        }

        return $this->redirectToRoute('dossier_show', ['id' => $parent->getId()]);
    }

    /** @return array<string, mixed> */
    private function config(string $slug): array
    {
        return Catalogue::all()[$slug] ?? throw $this->createNotFoundException();
    }
}
