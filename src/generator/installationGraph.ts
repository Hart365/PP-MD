import type { MissingDependency, ParsedSolution } from '../types/solution';
import { dependencyKey } from './dependencyAnalysis';

export interface InstallationPrerequisiteEdge {
  consumer: ParsedSolution;
  provider: ParsedSolution;
  reason: string;
  dependency?: MissingDependency;
}

export interface InstallationBlocker {
  solution: ParsedSolution;
  kind: 'direct' | 'cycle' | 'waiting' | 'undetermined';
  reason: string;
  dependency?: MissingDependency;
  detail?: string;
}

export interface InstallationCycle {
  id: string;
  members: ParsedSolution[];
  edges: InstallationPrerequisiteEdge[];
  waiting: ParsedSolution[];
}

type Prerequisites = Map<ParsedSolution, Map<ParsedSolution, Set<string>>>;

export function buildInstallationGraph(
  solutions: ParsedSolution[],
  prerequisites: Prerequisites,
  edges: InstallationPrerequisiteEdge[],
  installable: ParsedSolution[],
) {
  const sorted = [...solutions].sort((left, right) =>
    left.metadata.uniqueName.localeCompare(right.metadata.uniqueName)
    || left.metadata.version.localeCompare(right.metadata.version)
    || Number(left.metadata.isManaged) - Number(right.metadata.isManaged));
  const ids = new Map(sorted.map((solution) => [
    solution,
    `solution:${encodeURIComponent(dependencyKey(solution.metadata.uniqueName))}:${encodeURIComponent(solution.metadata.version)}:${String(solution.metadata.isManaged)}`,
  ]));
  const consumers = new Map(solutions.map((solution) => [solution, new Set<ParsedSolution>()]));
  prerequisites.forEach((providers, consumer) => providers.forEach((_, provider) => consumers.get(provider)?.add(consumer)));

  const indexes = new Map<ParsedSolution, number>();
  const lowLinks = new Map<ParsedSolution, number>();
  const stack: ParsedSolution[] = [];
  const onStack = new Set<ParsedSolution>();
  const components: ParsedSolution[][] = [];
  let nextIndex = 0;
  const visit = (solution: ParsedSolution) => {
    const index = nextIndex++;
    indexes.set(solution, index);
    lowLinks.set(solution, index);
    stack.push(solution);
    onStack.add(solution);
    for (const provider of prerequisites.get(solution)?.keys() ?? []) {
      if (!indexes.has(provider)) {
        visit(provider);
        lowLinks.set(solution, Math.min(lowLinks.get(solution)!, lowLinks.get(provider)!));
      } else if (onStack.has(provider)) {
        lowLinks.set(solution, Math.min(lowLinks.get(solution)!, indexes.get(provider)!));
      }
    }
    if (lowLinks.get(solution) !== indexes.get(solution)) return;
    const component: ParsedSolution[] = [];
    let member: ParsedSolution;
    do {
      member = stack.pop()!;
      onStack.delete(member);
      component.push(member);
    } while (member !== solution);
    if (component.length > 1) components.push(component);
  };
  sorted.forEach((solution) => { if (!indexes.has(solution)) visit(solution); });
  components.forEach((component) => component.sort((left, right) => ids.get(left)!.localeCompare(ids.get(right)!)));
  components.sort((left, right) => ids.get(left[0])!.localeCompare(ids.get(right[0])!));

  const cycles: InstallationCycle[] = components.map((members, index) => {
    const memberSet = new Set(members);
    const affected = new Set(members);
    const queue = [...members];
    for (let cursor = 0; cursor < queue.length; cursor += 1) {
      consumers.get(queue[cursor])?.forEach((consumer) => {
        if (!affected.has(consumer)) {
          affected.add(consumer);
          queue.push(consumer);
        }
      });
    }
    return {
      id: `Cycle ${index + 1}`,
      members,
      edges: edges.filter((edge) => memberSet.has(edge.consumer) && memberSet.has(edge.provider)),
      waiting: sorted.filter((solution) => affected.has(solution) && !memberSet.has(solution)),
    };
  });

  const readiness = new Map<ParsedSolution, number>();
  installable.forEach((solution) => {
    const providerLayers = [...(prerequisites.get(solution)?.keys() ?? [])].map((provider) => {
      const layer = readiness.get(provider);
      if (layer === undefined) throw new Error(`Installation analysis placed ${solution.metadata.uniqueName} before prerequisite ${provider.metadata.uniqueName}.`);
      return layer;
    });
    readiness.set(solution, providerLayers.length ? Math.max(...providerLayers) + 1 : 0);
  });
  return { ids, cycles, readiness };
}
