import type { MissingDependency, ParsedSolution } from '../types/solution';
import { missingDependenciesNotice } from './missingDependencies';
import {
  analyzeDependencies, dependencyKey, describeDependencyLocations, describeDependent, describeRequired,
  requiredComponentIdentity, solutionNameOf, DEPENDENCY_STATUS_LABELS,
  type DependencyResolution, type DependencyStatus,
} from './dependencyAnalysis';
import { buildInstallationGraph, type InstallationBlocker, type InstallationPrerequisiteEdge } from './installationGraph';

const PLATFORM_SOLUTIONS = /^(system|basic|active|default)$/i;
const MICROSOFT_SOLUTIONS = /^(ms[a-z0-9]*_|msdyn|mscrm|msft|microsoft)/i;
const unique = (values: string[]) => [...new Set(values.filter(Boolean))];

function compareVersions(left: string, right: string): number {
  const a = left.split('.').map(Number);
  const b = right.split('.').map(Number);
  for (let index = 0; index < Math.max(a.length, b.length); index += 1) {
    const difference = (a[index] ?? 0) - (b[index] ?? 0);
    if (difference) return difference;
  }
  return 0;
}

export function analyzeInstallation(solutions: ParsedSolution[]) {
  const analysis = analyzeDependencies(solutions);
  const selected = analysis.solutions;
  const name = (solution: ParsedSolution) => solution.metadata.uniqueName;
  const prerequisites = new Map(selected.map((solution) => [solution, new Map<ParsedSolution, Set<string>>()]));
  const prerequisiteEdges: InstallationPrerequisiteEdge[] = [];
  const blockers = new Map<ParsedSolution, string[]>();
  const blockerDetails: InstallationBlocker[] = [];
  const missingRows: string[][] = [];
  const environmentRows: string[][] = [];
  const targetPrerequisites: Array<{ consumer: ParsedSolution; owner: string; required: string; detail: string; dependency?: MissingDependency }> = [];
  const templateRows: string[][] = [];
  const inferredReferences: Array<{ consumer: ParsedSolution; dependency: MissingDependency; status: DependencyStatus; providers: ParsedSolution[]; resolution: string }> = [];
  const notes = new Set<string>();

  const block = (solution: ParsedSolution, reason: string, kind: InstallationBlocker['kind'] = 'direct', dependency?: MissingDependency, detail?: string) => {
    blockers.set(solution, unique([...(blockers.get(solution) ?? []), reason]));
    if (!blockerDetails.some((entry) => entry.solution === solution && entry.reason === reason && entry.kind === kind && entry.dependency === dependency)) {
      blockerDetails.push({ solution, reason, kind, dependency, detail });
    }
  };
  const addIssue = (consumer: string, dependent: string, required: string, owner: string, found: string, action: string) =>
    missingRows.push([consumer, dependent, required, '–', '–', '–', '–', owner, found, action]);
  const require = (solution: ParsedSolution, provider: ParsedSolution, reason: string, dependency?: MissingDependency) => {
    if (provider === solution) return;
    const reasons = prerequisites.get(solution);
    if (reasons) reasons.set(provider, new Set([...(reasons.get(provider) ?? []), reason]));
    prerequisiteEdges.push({ consumer: solution, provider, reason, dependency });
  };
  const reaches = (source: ParsedSolution, target: ParsedSolution, visited = new Set<ParsedSolution>()): boolean => {
    if (source === target) return true;
    if (visited.has(source)) return false;
    visited.add(source);
    return [...(prerequisites.get(source)?.keys() ?? [])].some((provider) => reaches(provider, target, visited));
  };

  selected.forEach((solution) => {
    const variants = selected.filter((candidate) => dependencyKey(candidate.metadata.uniqueName) === dependencyKey(solution.metadata.uniqueName));
    if (variants.length > 1) {
      const reason = `Several ZIPs of ${name(solution)} are selected (${variants.map((item) => item.metadata.version).join(', ')}). Select only one.`;
      block(solution, reason);
      addIssue(name(solution), 'Solution import', name(solution), 'Duplicate selection', 'Selected ZIPs', reason);
    }
    const notice = missingDependenciesNotice(solution);
    if (notice) {
      block(solution, 'Dependency manifest unavailable or incomplete.');
      addIssue(name(solution), 'Solution import', 'Dependency manifest', '–', 'ZIP manifest', notice);
    }
    solution.warnings.filter((warning) => /^solution\.xml contains .*dependency entr.*without a solution name/.test(warning)).forEach((warning) => {
      block(solution, 'Dependency manifest contains unreadable entries.');
      addIssue(name(solution), 'Solution import', 'Dependency manifest', '–', 'ZIP manifest', warning);
    });
    solution.metadata.dependencies.filter((dependency) => !dependency.isInternal && dependency.versionSource !== 'exported').forEach((dependency) => {
      const required = solutionNameOf(dependency.solutionName);
      const providers = analysis.solutionsNamed(required).filter((provider) => provider !== solution);
      const minimum = dependency.version ? ` (minimum ${dependency.version})` : '';
      if (providers.length === 1) {
        require(solution, providers[0], 'Declared solution dependency');
        if (dependency.version && (!/^\d+(\.\d+){0,3}$/.test(providers[0].metadata.version)
          || !/^\d+(\.\d+){0,3}$/.test(dependency.version)
          || compareVersions(providers[0].metadata.version, dependency.version) < 0)) {
          const reason = `Requires ${required} ${dependency.version} or later; the selected ZIP is ${providers[0].metadata.version}.`;
          block(solution, reason);
          addIssue(name(solution), 'Solution import', `Solution: ${required}${minimum}`, required, 'Selected ZIP', `${reason} Select a compatible version.`);
        }
      } else if (PLATFORM_SOLUTIONS.test(required) || MICROSOFT_SOLUTIONS.test(required)) {
        const owner = `Solution: ${required}${minimum}`;
        const detail = 'Microsoft/platform solution. Verify it is installed in the target environment.';
        environmentRows.push([owner, 'Solution import', name(solution), detail]);
        targetPrerequisites.push({ consumer: solution, owner, required: 'Solution import', detail });
      } else {
        block(solution, `Requires solution ${required}, which is not selected.`);
        addIssue(name(solution), 'Solution import', `Solution: ${required}${minimum}`, required, 'Not in selected ZIPs',
          'Install it in the target environment first, or add its ZIP and regenerate.');
      }
    });
  });

  analysis.resolutions.forEach((resolution) => {
    if (resolution.status === 'selected') resolution.providers.forEach((provider) =>
      require(resolution.consumerSolution, provider, 'Import dependency', resolution.dependency));
    if (resolution.status === 'missing') {
      block(resolution.consumerSolution, 'Missing import dependencies.', 'direct', resolution.dependency, resolution.reason);
      const identity = requiredComponentIdentity(resolution.dependency);
      missingRows.push([
        resolution.consumer, describeDependent(resolution.dependency), describeRequired(resolution.dependency),
        identity.table || '–', identity.column || '–', identity.schemaName || 'Unknown / not supplied',
        identity.id || 'Not supplied', resolution.owner ?? '–', describeDependencyLocations(resolution.dependency), resolution.reason,
      ]);
    }
    if (resolution.status === 'environment') {
      const owner = resolution.owner ?? describeRequired(resolution.dependency);
      const required = describeRequired(resolution.dependency);
      environmentRows.push([owner, required, resolution.consumer, resolution.reason]);
      targetPrerequisites.push({ consumer: resolution.consumerSolution, owner, required, detail: resolution.reason, dependency: resolution.dependency });
    }
  });

  selected.forEach((solution) => solution.emailTemplates.forEach((template) => {
    const references = [
      ...(template.entityLogicalName && template.entityLogicalName !== 'global' && !/^\d+$/.test(template.entityLogicalName)
        ? [{ requiredType: 'Table', requiredName: template.entityLogicalName }] : []),
      ...(template.referencedFields ?? []).flatMap((field) => [
        { requiredType: 'Table', requiredName: field.entityLogicalName },
        { requiredType: 'Column', requiredName: field.attributeLogicalName, requiredParent: field.entityLogicalName },
      ]),
    ];
    references.forEach((reference) => {
      const dependency: MissingDependency = { ...reference, dependentName: template.name, dependentType: 'Email Template', tableLogicalNames: [] };
      const result = analysis.classify(solution, dependency);
      let resolution = result.reason;
      if (result.status === 'selected') {
        const conflicts = result.providers.filter((provider) => reaches(provider, solution));
        if (conflicts.length) {
          block(solution, 'Inferred email template reference conflicts with the declared order.', 'direct', dependency);
          resolution = `Inferred reference to ${conflicts.map(name).join(', ')} conflicts with the declared order; manual review required.`;
        } else {
          result.providers.forEach((provider) => require(solution, provider, 'Inferred email template reference', dependency));
          resolution = `Install after ${result.providers.map(name).join(', ')} (inferred, not a declared prerequisite).`;
        }
      } else if (result.status === 'missing') {
        resolution = 'Not found in selected ZIPs or the standard platform. Verify it in the target environment; this inferred reference is not a blocking import prerequisite.';
        notes.add(`Verify inferred ${reference.requiredType.toLowerCase()} ${[('requiredParent' in reference ? reference.requiredParent : ''), reference.requiredName].filter(Boolean).join('.')} used by ${name(solution)} email templates.`);
      }
      templateRows.push([name(solution), template.name, describeRequired(dependency), resolution]);
      inferredReferences.push({ consumer: solution, dependency, status: result.status, providers: result.providers, resolution });
    });
  }));

  const ordered: ParsedSolution[] = [];
  const remaining = new Set(selected);
  while (remaining.size) {
    const ready = [...remaining].filter((solution) =>
      [...(prerequisites.get(solution)?.keys() ?? [])].every((provider) => !remaining.has(provider)));
    if (!ready.length) break;
    ready.sort((left, right) => name(left).localeCompare(name(right)));
    ready.forEach((solution) => { ordered.push(solution); remaining.delete(solution); });
  }
  const graphCycles = buildInstallationGraph(selected, prerequisites, prerequisiteEdges, []);
  const cycleMembers = new Set(graphCycles.cycles.flatMap((cycle) => cycle.members));
  graphCycles.cycles.forEach((cycle) => cycle.members.forEach((solution) =>
    block(solution, `Circular dependency ${cycle.id}: no safe sequential installation order satisfies all members.`, 'cycle')));
  const remainingBlocked = new Set([...remaining].filter((solution) => !cycleMembers.has(solution)));
  let changed = true;
  while (changed) {
    changed = false;
    remainingBlocked.forEach((solution) => {
      const failed = [...(prerequisites.get(solution)?.keys() ?? [])].filter((provider) => blockers.has(provider));
      if (failed.length) {
        block(solution, `Waits for blocked prerequisite ${failed.map(name).join(', ')}.`, 'waiting');
        remainingBlocked.delete(solution);
        changed = true;
      }
    });
  }
  remainingBlocked.forEach((solution) =>
    block(solution, 'No safe installation order could be determined from the selected dependency graph.', 'undetermined'));
  ordered.forEach((solution) => {
    const failed = [...(prerequisites.get(solution)?.keys() ?? [])].filter((provider) => blockers.has(provider));
    if (failed.length) block(solution, `Waits for blocked prerequisite ${failed.map(name).join(', ')}.`, 'waiting');
  });
  const installable = ordered.filter((solution) => !blockers.has(solution));
  const graph = buildInstallationGraph(selected, prerequisites, prerequisiteEdges, installable);
  const circularRows = graph.cycles.flatMap((cycle) => cycle.edges.map((edge) => {
    const identity = edge.dependency ? requiredComponentIdentity(edge.dependency) : undefined;
    return [
      cycle.id, name(edge.consumer), name(edge.provider), edge.dependency ? describeDependent(edge.dependency) : 'Solution import',
      edge.dependency ? describeRequired(edge.dependency) : `Solution: ${name(edge.provider)}`,
      identity?.schemaName || '–', identity?.id || 'Not supplied', edge.reason,
    ];
  }));
  const stepOf = (solution: ParsedSolution) => {
    const index = installable.indexOf(solution);
    return index < 0 ? undefined : index + 1;
  };
  const stepLabel = (solution: ParsedSolution) => {
    const step = stepOf(solution);
    return step ? `Step ${step}` : 'Blocked';
  };
  const count = (status: DependencyStatus) => analysis.resolutions.filter((resolution) => resolution.status === status).length;
  const group = (status: DependencyStatus) => {
    const groups = new Map<string, DependencyResolution[]>();
    analysis.resolutions.filter((resolution) => resolution.status === status).forEach((resolution) => {
      const key = `${resolution.consumer}|${resolution.providers.map(name).sort().join(', ')}`;
      groups.set(key, [...(groups.get(key) ?? []), resolution]);
    });
    return [...groups.values()];
  };
  const selectedPairs = new Map<string, { consumer: ParsedSolution; provider: ParsedSolution; items: DependencyResolution[] }>();
  analysis.resolutions.filter((resolution) => resolution.status === 'selected').forEach((resolution) =>
    resolution.providers.filter((provider) => provider !== resolution.consumerSolution).forEach((provider) => {
      const key = `${name(resolution.consumerSolution)}|${name(provider)}`;
      const entry = selectedPairs.get(key) ?? { consumer: resolution.consumerSolution, provider, items: [] };
      entry.items.push(resolution);
      selectedPairs.set(key, entry);
    }));
  const selectedDependencyRows = [...selectedPairs.values()].map(({ consumer, provider, items }) => {
    const providerStep = stepOf(provider);
    const consumerStep = stepOf(consumer);
    const circular = graph.cycles.some((cycle) => cycle.members.includes(consumer) && cycle.members.includes(provider)
      && cycle.edges.some((edge) => edge.consumer === consumer && edge.provider === provider));
    const blocking = circular || !providerStep;
    const status = circular ? `Blocker: circular dependency between ${name(provider)} and ${name(consumer)}.`
      : !providerStep ? `Blocker: provider ${name(provider)} has no installation step.`
        : `Resolved by order: ${name(provider)} (step ${providerStep}) precedes ${name(consumer)} (${consumerStep ? `step ${consumerStep}` : 'blocked for another reason'}).`;
    return {
      blocking, row: [status, name(consumer), stepLabel(consumer), name(provider), stepLabel(provider),
        unique(items.map((item) => describeRequired(item.dependency))).join('; '), unique(items.map((item) => item.reason)).join(' ')],
    };
  }).sort((left, right) => Number(right.blocking) - Number(left.blocking)
    || left.row[1].localeCompare(right.row[1]) || left.row[3].localeCompare(right.row[3]));
  const environmentGroups = new Map<string, string[][]>();
  environmentRows.forEach((row) => environmentGroups.set(row[0], [...(environmentGroups.get(row[0]) ?? []), row]));
  const environmentTable = [...environmentGroups.entries()].map(([owner, rows]) =>
    [owner, unique(rows.map((row) => row[1])).join('; '), unique(rows.map((row) => row[2])).join(', '), unique(rows.map((row) => row[3])).join(' ')]);
  analysis.notes.forEach((note) => notes.add(note));

  return {
    analysis, selected, prerequisites, prerequisiteEdges, blockers, blockerDetails, cycleMembers, installable,
    ids: graph.ids, cycles: graph.cycles, readiness: graph.readiness,
    missingRows, environmentTable, targetPrerequisites, templateRows, inferredReferences, notes, circularRows,
    selectedDependencyRows,
    blockingSelectedCount: selectedDependencyRows.filter((entry) => entry.blocking).length,
    resolvedRows: selectedDependencyRows.map((entry) => entry.row),
    includedRows: group('included').map((items) => [items[0].consumer, String(unique(items.map((item) => describeRequired(item.dependency))).length),
      unique(items.map((item) => describeRequired(item.dependency))).join('; ')]),
    platformRows: group('platform').map((items) => [items[0].consumer, unique(items.map((item) => describeRequired(item.dependency))).join('; ')]),
    searchRows: analysis.resolutions.map(({ consumer, dependency, status }) => [
      consumer, describeDependent(dependency), dependency.dependentId || 'Not supplied',
      describeRequired(dependency), dependency.requiredId || 'Not supplied', DEPENDENCY_STATUS_LABELS[status],
      describeDependencyLocations(dependency),
    ]),
    count, stepOf, stepLabel,
  };
}

export type InstallationAnalysis = ReturnType<typeof analyzeInstallation>;
