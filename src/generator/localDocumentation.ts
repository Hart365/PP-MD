import type { MissingDependency, ParsedSolution } from '../types/solution';
import { missingDependenciesNotice, MISSING_DEPENDENCIES_NOTE } from './missingDependencies';
import { heading, mdEscape } from './markdownFormatting';
import {
  analyzeDependencies, dependencyKey, describeDependent, describeRequired, solutionNameOf,
  describeDependencyLocations, requiredComponentIdentity, type DependencyResolution, type DependencyStatus,
} from './dependencyAnalysis';
import { analyzeInstallation, type InstallationAnalysis } from './installationAnalysis';

const PLATFORM_SOLUTIONS = /^(system|basic|active|default)$/i;
const MICROSOFT_SOLUTIONS = /^(ms[a-z0-9]*_|msdyn|mscrm|msft|microsoft)/i;

const table = (headers: string[], rows: string[][]) => [
  `| ${headers.join(' | ')} |`, `| ${headers.map(() => '---').join(' | ')} |`,
  ...[...new Set(rows.map((row) => `| ${row.map(mdEscape).join(' | ')} |`))],
];

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

/** @deprecated Compatibility helper; application generation uses the reusable installation planner below. */
export function generateLegacyInstallationMarkdown(solutions: ParsedSolution[]): string {
  const analysis = analyzeDependencies(solutions);
  const selected = analysis.solutions;
  const name = (solution: ParsedSolution) => solution.metadata.uniqueName;
  const prerequisites = new Map(selected.map((solution) => [solution, new Map<ParsedSolution, Set<string>>()]));
  const blockers = new Map<ParsedSolution, string[]>();
  const missingRows: string[][] = [];
  const environmentRows: string[][] = [];
  const templateRows: string[][] = [];
  const notes = new Set<string>();
  const circularRows: string[][] = [];
  const block = (solution: ParsedSolution, reason: string) => blockers.set(solution, unique([...(blockers.get(solution) ?? []), reason]));
  const require = (solution: ParsedSolution, provider: ParsedSolution, reason: string) => {
    if (provider === solution) return;
    const reasons = prerequisites.get(solution)!;
    reasons.set(provider, new Set([...(reasons.get(provider) ?? []), reason]));
  };
  const reaches = (source: ParsedSolution, target: ParsedSolution, visited = new Set<ParsedSolution>()): boolean => {
    if (source === target) return true;
    if (visited.has(source)) return false;
    visited.add(source);
    return [...prerequisites.get(source)!.keys()].some((provider) => reaches(provider, target, visited));
  };

  selected.forEach((solution) => {
    const versions = selected.filter((candidate) => dependencyKey(name(candidate)) === dependencyKey(name(solution)));
    if (versions.length > 1) {
      const reason = `Several ZIPs of ${name(solution)} are selected (${versions.map((item) => item.metadata.version).join(', ')}). Select only one.`;
      block(solution, reason);
      missingRows.push([name(solution), 'Solution import', name(solution), '', '', '', '', 'Duplicate selection', 'Solution selection', reason]);
    }
    const notice = missingDependenciesNotice(solution);
    if (notice) {
      block(solution, 'Dependency manifest unavailable or incomplete.');
      missingRows.push([name(solution), 'Solution import', 'Dependency manifest', '', '', '', '', '–', 'ZIP manifest', notice]);
    }
    solution.warnings.filter((warning) => /^solution\.xml contains .*dependency entr.*without a solution name/.test(warning)).forEach((warning) => {
      block(solution, 'Dependency manifest contains unreadable entries.');
      missingRows.push([name(solution), 'Solution import', 'Dependency manifest', '', '', '', '', '–', 'ZIP manifest', warning]);
    });
    solution.metadata.dependencies.filter((dependency) => !dependency.isInternal && dependency.versionSource !== 'exported').forEach((dependency) => {
      const required = solutionNameOf(dependency.solutionName);
      const providers = analysis.solutionsNamed(required).filter((provider) => provider !== solution);
      const minimum = dependency.version ? ` (minimum ${dependency.version})` : '';
      if (providers.length === 1) {
        require(solution, providers[0], 'Declared solution dependency');
        const actual = providers[0].metadata.version;
        if (dependency.version && (!/^\d+(\.\d+)*$/.test(actual) || !/^\d+(\.\d+)*$/.test(dependency.version) || compareVersions(actual, dependency.version) < 0)) {
          const reason = `Requires ${required} ${dependency.version} or later; the selected ZIP is ${actual}.`;
          block(solution, reason);
          missingRows.push([name(solution), 'Solution import', `Solution: ${required}${minimum}`, '', '', '', '', required, 'Declared dependency', `${reason} Select a compatible version.`]);
        }
      } else if (PLATFORM_SOLUTIONS.test(required) || MICROSOFT_SOLUTIONS.test(required)) {
        environmentRows.push([`Solution: ${required}${minimum}`, 'Solution import', name(solution), 'Microsoft/platform solution. Verify it is installed in the target environment.']);
      } else {
        block(solution, `Requires solution ${required}, which is not selected.`);
        missingRows.push([name(solution), 'Solution import', `Solution: ${required}${minimum}`, '', '', '', '', required, 'Declared dependency', 'Not among the selected ZIPs. Install it in the target environment first, or add its ZIP and regenerate.']);
      }
    });
  });

  analysis.resolutions.forEach((resolution) => {
    if (resolution.status === 'selected') resolution.providers.forEach((provider) => require(resolution.consumerSolution, provider, 'Import dependency'));
    if (resolution.status === 'missing') {
      block(resolution.consumerSolution, 'Missing import dependencies.');
      const identity = requiredComponentIdentity(resolution.dependency);
      missingRows.push([
        resolution.consumer, describeDependent(resolution.dependency), describeRequired(resolution.dependency),
        identity.table || '–', identity.column || '–', identity.schemaName || 'Unknown / not supplied',
        identity.id || 'Not supplied', resolution.owner ?? '–',
        describeDependencyLocations(resolution.dependency), resolution.reason,
      ]);
    }
    if (resolution.status === 'environment') {
      environmentRows.push([resolution.owner ?? describeRequired(resolution.dependency), describeRequired(resolution.dependency), resolution.consumer, resolution.reason]);
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
      let resolution: string = result.reason;
      if (result.status === 'selected') {
        const conflicts = result.providers.filter((provider) => reaches(provider, solution));
        if (conflicts.length) {
          block(solution, 'Inferred email template reference conflicts with the declared order.');
          resolution = `Inferred reference to ${conflicts.map(name).join(', ')} conflicts with the declared order; manual review required.`;
        } else {
          result.providers.forEach((provider) => require(solution, provider, 'Inferred email template reference'));
          resolution = `Install after ${result.providers.map(name).join(', ')} (inferred from the template, not a declared import prerequisite).`;
        }
      } else if (result.status === 'missing') {
        resolution = 'Not found in the selected ZIPs or the standard platform. Verify it exists in the target environment (inferred reference; not treated as a missing import prerequisite).';
        notes.add(`Verify inferred ${reference.requiredType.toLowerCase()} ${[('requiredParent' in reference ? reference.requiredParent : ''), reference.requiredName].filter(Boolean).join('.')} used by ${name(solution)} email templates.`);
      }
      templateRows.push([name(solution), template.name, describeRequired(dependency), resolution]);
    });
  }));

  const ordered: ParsedSolution[] = [];
  const remaining = new Set(selected);
  while (remaining.size) {
    const ready = [...remaining].filter((solution) => [...prerequisites.get(solution)!.keys()].every((provider) => !remaining.has(provider)));
    if (!ready.length) break;
    ready.forEach((solution) => { ordered.push(solution); remaining.delete(solution); });
  }
  const cycleMembers = new Set([...remaining].filter((solution) =>
    [...prerequisites.get(solution)!.keys()].some((provider) =>
      remaining.has(provider) && reaches(provider, solution))));
  cycleMembers.forEach((solution) => {
    const cycle = [...prerequisites.get(solution)!.entries()]
      .filter(([provider]) => cycleMembers.has(provider) && reaches(provider, solution));
    cycle.forEach(([provider, reasons]) => {
      circularRows.push([name(solution), name(provider), 'Solution import', `Solution: ${name(provider)}`,
        [...reasons].join(', ')]);
    });
    block(solution, `Circular dependency: requires ${cycle.map(([provider, reasons]) =>
      `${name(provider)} (${[...reasons].join(', ')})`).join('; ') || 'another member of the cycle'}, which in turn requires this solution. No safe installation order exists.`);
  });
  const waitingSolutions = new Set([...remaining].filter((solution) => !cycleMembers.has(solution)));
  let changed = true;
  while (changed) {
    changed = false;
    waitingSolutions.forEach((solution) => {
      const failed = [...prerequisites.get(solution)!.keys()].filter((provider) =>
        cycleMembers.has(provider) || blockers.has(provider));
      if (failed.length) {
        block(solution, `Waits for blocked prerequisite ${failed.map(name).join(', ')}.`);
        waitingSolutions.delete(solution);
        changed = true;
      }
    });
  }
  waitingSolutions.forEach((solution) =>
    block(solution, 'No safe installation order could be determined from the selected dependency graph.'));
  ordered.forEach((solution) => {
    const failed = [...prerequisites.get(solution)!.keys()].filter((provider) => blockers.has(provider));
    if (failed.length) block(solution, `Waits for blocked prerequisite ${failed.map(name).join(', ')}.`);
  });
  const installable = ordered.filter((solution) => !blockers.has(solution));

  const count = (status: DependencyStatus) => analysis.resolutions.filter((resolution) => resolution.status === status).length;
  const group = (status: DependencyStatus) => {
    const groups = new Map<string, DependencyResolution[]>();
    analysis.resolutions.filter((resolution) => resolution.status === status).forEach((resolution) => {
      const key = `${resolution.consumer}|${resolution.providers.map(name).sort().join(', ')}`;
      groups.set(key, [...(groups.get(key) ?? []), resolution]);
    });
    return [...groups.values()];
  };
  const resolvedRows = group('selected').map((items) => [items[0].consumer, items[0].providers.map(name).sort().join(', '),
    unique(items.map((item) => describeRequired(item.dependency))).join('; '), unique(items.map((item) => item.reason)).join(' ')]);
  const includedRows = group('included').map((items) => [items[0].consumer, String(unique(items.map((item) => describeRequired(item.dependency))).length),
    unique(items.map((item) => describeRequired(item.dependency))).join('; ')]);
  const platformRows = group('platform').map((items) => [items[0].consumer, unique(items.map((item) => describeRequired(item.dependency))).join('; ')]);
  const environmentGroups = new Map<string, string[][]>();
  environmentRows.forEach((row) => environmentGroups.set(row[0], [...(environmentGroups.get(row[0]) ?? []), row]));
  const environmentTable = [...environmentGroups.entries()].map(([owner, rows]) => [owner,
    unique(rows.map((row) => row[1])).join('; '), unique(rows.map((row) => row[2])).join(', '), unique(rows.map((row) => row[3])).join(' ')]);
  analysis.notes.forEach((note) => notes.add(note));
  const missingCount = missingRows.length;

  return [heading(1, 'Solution Installation Guide'), '',
    'This guide analyses the selected solution ZIPs offline. Each declared import prerequisite (the ZIP MissingDependencies manifest) is checked against the solution that declares it, the other selected ZIPs, the standard Dataverse platform and Microsoft-supplied solutions. It does not query the target environment, so verify environment prerequisites before importing.', '',
    heading(2, 'Summary'), '',
    ...table(['Measure', 'Count'], [
      ['Selected solutions', String(selected.length)],
      ['Ready to install (numbered steps)', String(installable.length)],
      ['Blocked solutions', String(blockers.size)],
      ['Declared import prerequisites', String(analysis.resolutions.length)],
      ['Included in the declaring solution', String(count('included'))],
      ['Provided by another selected solution', String(count('selected'))],
      ['Standard platform components', String(count('platform'))],
      ['Target environment prerequisites to verify', String(environmentTable.length)],
      ['Missing dependencies', String(missingCount)],
    ]), '',
    heading(2, 'Installation Order'), '',
    blockers.size ? '**Action required: blocked solutions are not installation steps. Resolve the Missing Dependencies below and regenerate this guide.**'
      : 'No missing dependencies were found. Import the solutions in this order, waiting for each import to finish before starting the next.', '',
    ...(installable.length ? table(['Order', 'Solution', 'Version', 'Install After', 'Status'], installable.map((solution, index) => [
      String(index + 1), name(solution), solution.metadata.version,
      [...prerequisites.get(solution)!.keys()].map(name).sort().join(', ') || 'No selected prerequisites', 'Ready',
    ])) : ['No solution can be installed until the issues below are resolved.']), '',
    ...(blockers.size ? [heading(2, 'Blocked Solutions'), '', ...table(['Solution', 'Reason'],
      selected.filter((solution) => blockers.has(solution)).map((solution) => [name(solution), blockers.get(solution)!.join(' ')])), ''] : []),
    ...(circularRows.length ? [heading(2, 'Circular Dependencies'), '',
      'These solution-to-solution edges form cycles; no sequential import order can satisfy them. Split shared components into a base solution or remove the cyclic references, then regenerate. Solutions that only depend on a cycle appear as waiting for a blocked prerequisite.', '',
      ...table(['Solution', 'Depends On', 'Dependent Component', 'Required Component', 'Evidence'], circularRows), ''] : []),
    heading(2, 'Missing Dependencies'), '',
    ...(missingRows.length ? [
      'These prerequisites are not in any selected ZIP and are not standard or Microsoft-supplied components. The dependent solution cannot be imported until they exist in the target environment.', '',
      ...table(['Dependent Solution', 'Dependent Component', 'Required Component', 'Table', 'Column', 'Schema / Logical Name', 'Component ID', 'Expected Owner', 'Found in', 'Action'], missingRows),
    ] : ['No missing dependencies. Every declared prerequisite is included in its own solution, provided by another selected solution, or is a standard/environment component.']), '',
    ...(environmentTable.length ? [heading(2, 'Target Environment Prerequisites'), '',
      'Supplied by Microsoft or installed applications rather than the selected ZIPs. They are not missing from your selection, but they must exist in the target environment.', '',
      ...table(['Owner / Application', 'Required Components', 'Needed By', 'Action'], environmentTable), ''] : []),
    ...(resolvedRows.length ? [heading(2, 'Dependencies Provided by Selected Solutions'), '',
      ...table(['Solution', 'Install After', 'Required Components', 'Evidence'], resolvedRows), ''] : []),
    ...(includedRows.length ? [heading(2, 'Dependencies Included in the Same Solution'), '',
      'Declared prerequisites that are defined in the declaring solution ZIP itself. No action is needed.', '',
      ...table(['Solution', 'Count', 'Components'], includedRows), ''] : []),
    ...(platformRows.length ? [heading(2, 'Standard Platform Components'), '',
      'Standard Dataverse tables, columns and choices that exist in every environment. No action is needed.', '',
      ...table(['Solution', 'Components'], platformRows), ''] : []),
    ...(templateRows.length ? [heading(2, 'Inferred Email Template References'), '',
      'Derived from email template metadata and placeholders, not the MissingDependencies manifest. Selected providers influence the order; unresolved references are verification items, not missing dependencies.', '',
      ...table(['Solution', 'Email Template', 'Referenced Component', 'Resolution'], templateRows), ''] : []),
    ...(notes.size ? [heading(2, 'Verification Notes'), '', ...[...notes].sort().map((note) => `- ${mdEscape(note)}`), ''] : []),
    heading(2, 'How Dependencies Were Evaluated'), '',
    '1. **Included in the same solution:** the component ID, schema name or key is defined in the declaring ZIP (root components, customizations, workflows, web resources and other parsed components).',
    '2. **Provided by a selected solution:** the manifest names a selected solution as owner, or another selected ZIP defines the component. That solution is placed earlier in the order.',
    '3. **Standard platform component:** owned by System/Basic, or a standard table, column or choice name without a customization prefix.',
    '4. **Target environment prerequisite:** owned by a Microsoft managed solution or installed application package, or an unmanaged customization of a standard form/view.',
    '5. **Missing:** anything else, including components owned by unselected custom solutions or unmanaged custom components from the source environment that are not in any selected ZIP.', '',
    heading(2, 'Installation Checklist'), '',
    '1. Resolve every Missing Dependency: add the owning solution ZIP and regenerate, or install it in the target environment first.',
    '2. Confirm the Target Environment Prerequisites are installed and compatible.',
    '3. Back up the target environment and confirm publisher, managed/unmanaged import mode and solution versions.',
    '4. Import each numbered solution in order, waiting for each import to complete successfully.',
    '5. Configure connection references and environment variables, publish customizations, enable flows and verify apps and integrations.', '',
  ].join('\n');
}

export function renderInstallationMarkdown(plan: InstallationAnalysis): string {
  const {
    analysis, selected, prerequisites, blockers, installable, missingRows, environmentTable,
    templateRows, notes, circularRows, selectedDependencyRows, blockingSelectedCount,
    resolvedRows, includedRows, platformRows, searchRows, count, stepOf, stepLabel, readiness,
  } = plan;
  const name = (solution: ParsedSolution) => solution.metadata.uniqueName;
  const environmentCount = environmentTable.length;
  const installRows = installable.map((solution, index) => [
    String(index + 1), name(solution), solution.metadata.version,
    [...(prerequisites.get(solution)?.keys() ?? [])].sort((left, right) =>
      (stepOf(left) ?? 0) - (stepOf(right) ?? 0) || name(left).localeCompare(name(right)))
      .map(name).join(', ') || 'No selected prerequisites',
    'Ready',
  ]);
  const readinessRows = installable.map((solution) => [
    `Layer ${readiness.get(solution)}`, name(solution), solution.metadata.version, String(stepOf(solution)),
    [...(prerequisites.get(solution)?.keys() ?? [])].map((provider) =>
      `${name(provider)} (${stepLabel(provider).toLowerCase()})`).join(', ') || 'No selected prerequisites',
    'Ready',
  ]);
  const blockedRows = selected.filter((solution) => blockers.has(solution))
    .map((solution) => [name(solution), blockers.get(solution)!.join(' ')]);
  return [
    heading(1, 'Solution Installation Guide'), '',
    'This guide analyses only the Power Platform solution ZIPs selected by the user. It calculates a provider-first installation order from solution-level dependencies and declared `MissingDependencies`, then identifies standard platform items, target-environment checks, missing components, blockers, circular dependency groups and downstream solutions that wait for them. It does not query the target environment.', '',
    heading(2, 'Summary'), '',
    ...table(['Measure', 'Count'], [
      ['Selected solutions', String(selected.length)],
      ['Ready to install (numbered steps)', String(installable.length)],
      ['Blocked solutions', String(blockers.size)],
      ['Declared import prerequisites', String(analysis.resolutions.length)],
      ['Included in the declaring solution', String(count('included'))],
      ['Provided by another selected solution', String(count('selected'))],
      ['Selected-solution dependencies resolved by order', String(selectedDependencyRows.length - blockingSelectedCount)],
      ['Selected-solution dependencies that are blockers', String(blockingSelectedCount)],
      ['Standard platform components', String(count('platform'))],
      ['Target environment prerequisites to verify', String(environmentCount)],
      ['Blocking issues', String(missingRows.length)],
    ]), '',
    heading(2, 'Installation Order'), '',
    !selected.length ? 'No solutions were selected; there is no installation order to calculate.'
      : blockers.size ? '**Action required:** blocked solutions do not receive installation steps. Resolve the blockers below and regenerate the guide.'
        : 'No blocking dependencies were found. Import the numbered solutions sequentially and wait for each import to finish.', '',
    ...(installRows.length ? table(['Order', 'Solution', 'Version', 'Install After', 'Status'], installRows)
      : selected.length ? ['No solution can be installed until the issues below are resolved.'] : []), '',
    heading(2, 'Dependency Readiness Layers'), '',
    'Layers show prerequisite readiness only; they do not permit parallel imports. Follow the numbered order above.', '',
    ...(readinessRows.length ? table(['Layer', 'Solution', 'Version', 'Installation Step', 'Install After', 'Status'], readinessRows)
      : ['No selected solution currently has a safe numbered installation step.']), '',
    ...(blockedRows.length ? [heading(2, 'Blocked Solutions'), '',
      ...table(['Solution', 'Reason'], blockedRows), ''] : []),
    ...(circularRows.length ? [heading(2, 'Circular Dependencies'), '',
      'Only actual strongly connected cycle members appear in these edges. Downstream solutions that merely wait for a cycle are listed as waiting in Blocked Solutions, not as cycle members.', '',
      ...table(['Cycle', 'Solution', 'Depends On', 'Dependent Component', 'Required Component', 'Schema / Logical Name', 'Component ID', 'Evidence'], circularRows), ''] : []),
    heading(2, 'Missing Dependencies'), '',
    ...(missingRows.length ? [
      'These entries prevent a safe installation order. Add the component to a selected solution, install its package in the target environment, or select a compatible provider ZIP and regenerate.', '',
      ...table(['Dependent Solution', 'Dependent Component', 'Required Component', 'Table', 'Column', 'Schema / Logical Name', 'Component ID', 'Expected Owner', 'Found in', 'Action'], missingRows),
    ] : [selected.length
      ? 'No missing dependencies. Every declared prerequisite is included, provided by a selected solution, standard, or a target-environment prerequisite.'
      : 'No prerequisites were evaluated because no solutions were selected.']), '',
    ...(environmentTable.length ? [heading(2, 'Target Environment Prerequisites'), '',
      'These Microsoft/platform or application-delivered prerequisites are not missing from the selected set, but must exist and be compatible in the target environment.', '',
      ...table(['Owner / Application', 'Required Components', 'Needed By', 'Action'], environmentTable), ''] : []),
    ...(resolvedRows.length ? [heading(2, 'Dependencies Provided by Selected Solutions'), '',
      `${selectedDependencyRows.length - blockingSelectedCount} of ${selectedDependencyRows.length} selected-solution dependency pairs are resolved by installation order; blockers are listed first.`, '',
      ...table(['Status', 'Solution', 'Solution Step', 'Depends On', 'Provider Step', 'Required Components', 'Evidence'], resolvedRows), ''] : []),
    ...(includedRows.length ? [heading(2, 'Dependencies Included in the Same Solution'), '',
      ...table(['Solution', 'Count', 'Components'], includedRows), ''] : []),
    ...(platformRows.length ? [heading(2, 'Standard Platform Components'), '',
      ...table(['Solution', 'Components'], platformRows), ''] : []),
    ...(searchRows.length ? [heading(2, 'Dependency Evidence Audit'), '',
      'Every declared prerequisite is retained here, including resolved, standard and target-environment items.', '',
      ...table(['Selected Solution', 'Dependent Component', 'Dependent ID', 'Required Component', 'Required ID', 'Classification', 'Evidence'], searchRows), ''] : []),
    ...(templateRows.length ? [heading(2, 'Inferred Email Template References'), '',
      'These references are derived from template metadata/placeholders. They may affect provider ordering but an unresolved inferred reference is a verification item, not a missing declared import prerequisite.', '',
      ...table(['Solution', 'Email Template', 'Referenced Component', 'Resolution'], templateRows), ''] : []),
    ...(notes.size ? [heading(2, 'Verification Notes'), '', ...[...notes].sort().map((note) => `- ${mdEscape(note)}`), ''] : []),
    heading(2, 'How Dependencies Were Evaluated'), '',
    '1. **Included:** the component ID, schema name or key is defined in the declaring ZIP; table shells do not prove ownership of child columns.',
    '2. **Selected provider:** a selected ZIP declares/contains the required component. It must receive an earlier installation step.',
    '3. **Platform:** a recognised standard Dataverse table, column or choice.',
    '4. **Target environment prerequisite:** a Microsoft managed solution, installed application package or unmanaged source customisation that must be verified in the target.',
    '5. **Missing:** an undeclared/unselected component that cannot be satisfied by the standard platform or selected ZIPs.',
    '6. **Cycle:** actual strongly connected providers/consumers with no possible sequential order; downstream waiters are reported separately.', '',
    heading(2, 'Installation Checklist'), '',
    '1. Resolve every blocker and choose one compatible version of each solution.',
    '2. Confirm target-environment prerequisites, publisher mode and minimum versions.',
    '3. Back up the target environment and permissions before import.',
    '4. Import each numbered solution sequentially and wait for successful completion.',
    '5. Configure connections/environment variables, publish customisations, enable flows, and verify apps and integrations.', '',
  ].join('\n');
}

export function generateInstallationMarkdown(solutions: ParsedSolution[]): string {
  return renderInstallationMarkdown(analyzeInstallation(solutions));
}

export function generateDiagnosticsMarkdown(solution: ParsedSolution): string {
  const issues = [...new Set([
    ...solution.warnings,
    ...(solution.missingDependenciesError ? [solution.missingDependenciesError] : []),
    ...solution.environmentVariables.flatMap((variable) => variable.currentValueError ? [`${variable.schemaName}: ${variable.currentValueError}`] : []),
  ])];
  if (issues.length === 0) return '';
  return [heading(1, `${solution.metadata.displayName || solution.metadata.uniqueName}: Diagnostics`), '',
    heading(2, 'Parse Warnings and Issues'), '', ...issues.map((issue) => `- ${mdEscape(issue)}`), ''].join('\n');
}

export function resolveSolutionDependencies(solution: ParsedSolution, context: ParsedSolution[] = []): DependencyResolution[] {
  const scope = new Set((solution.dependencySolutionUniqueNames ?? [solution.metadata.uniqueName]).map(dependencyKey));
  const available = new Set(context.map((candidate) => dependencyKey(candidate.metadata.uniqueName)));
  const covered = context.length > 0 && [...scope].every((name) => available.has(name));
  const analysis = analyzeDependencies(covered ? context : [solution, ...context]);
  return analysis.resolutions.filter((resolution) => scope.has(dependencyKey(resolution.consumer)));
}

export function generateDependencyReportMarkdown(solution: ParsedSolution, context: ParsedSolution[] = [], level = 1): string {
  const resolutions = resolveSolutionDependencies(solution, context);
  const notice = missingDependenciesNotice(solution);
  const title = level === 1 ? `${solution.metadata.displayName || solution.metadata.uniqueName}: Import Dependencies` : 'Import Dependencies';
  const of = (status: DependencyStatus) => resolutions.filter((resolution) => resolution.status === status);
  const rows = (status: DependencyStatus) => of(status).map((resolution) => [
    resolution.consumer, resolution.dependency.dependentName, resolution.dependency.dependentId ?? '–', resolution.dependency.dependentType,
    describeRequired(resolution.dependency).replace(/^[^:]+: /, ''), resolution.dependency.requiredType,
    resolution.dependency.requiredId ?? '–',
    resolution.providers.map((provider) => provider.metadata.uniqueName).join(', ') || resolution.owner || '–',
    describeDependencyLocations(resolution.dependency), resolution.reason,
  ]);
  const headers = ['Solution', 'Dependent Item', 'Dependent ID', 'Item Type', 'Required Component', 'Required Type', 'Required ID', 'Owner / Provider', 'Found in', 'Detail'];
  const section = (status: DependencyStatus, sectionTitle: string, intro?: string) => of(status).length
    ? [heading(level + 1, sectionTitle), '', ...(intro ? [intro, ''] : []), ...table(headers, rows(status)), ''] : [];
  return [heading(level, title), '', MISSING_DEPENDENCIES_NOTE, '',
    ...(notice ? [notice, ''] : []),
    ...(resolutions.length ? [
      `${resolutions.length} declared prerequisites: ${of('missing').length} missing, ${of('environment').length} target environment prerequisites, ${of('platform').length} standard platform components, ${of('selected').length} provided by selected solutions, ${of('included').length} included in the same solution.`, '',
      heading(level + 1, 'Missing Dependencies'), '',
      ...(of('missing').length ? table(headers, rows('missing'))
        : ['No missing dependencies. Every declared prerequisite is included, provided by a selected solution, or a standard/environment component.']), '',
      ...section('environment', 'Target Environment Prerequisites', 'Not missing from the selection, but must exist in the target environment.'),
      ...section('selected', 'Provided by Selected Solutions'),
      ...section('included', 'Included in the Same Solution'),
      ...section('platform', 'Standard Platform Components'),
    ] : notice ? [] : ['No import prerequisites are declared in the solution ZIP.', '']),
  ].join('\n');
}
