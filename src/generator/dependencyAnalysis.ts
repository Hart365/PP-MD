import type { MissingDependency, ParsedSolution } from '../types/solution';
import { collectMissingDependencies } from './missingDependencies';

export type DependencyStatus = 'included' | 'selected' | 'platform' | 'environment' | 'missing';

export const DEPENDENCY_STATUS_LABELS: Record<DependencyStatus, string> = {
  included: 'Included in this solution',
  selected: 'Provided by selected solution',
  platform: 'Standard platform component',
  environment: 'Target environment prerequisite',
  missing: 'Missing',
};

export interface DependencyResolution {
  consumer: string;
  consumerSolution: ParsedSolution;
  dependency: MissingDependency;
  status: DependencyStatus;
  providers: ParsedSolution[];
  owner?: string;
  reason: string;
}

export interface DependencyAnalysis {
  solutions: ParsedSolution[];
  resolutions: DependencyResolution[];
  notes: string[];
  classify: (consumer: ParsedSolution, dependency: MissingDependency) => Omit<DependencyResolution, 'consumer' | 'consumerSolution' | 'dependency'>;
  solutionsNamed: (name: string) => ParsedSolution[];
}

export const dependencyKey = (value: string | undefined) => (value ?? '').trim().replace(/^\{|\}$/g, '').toLowerCase();
export const solutionNameOf = (value: string | undefined) => (value ?? '').replace(/\s*\([^)]*\)\s*$/, '').trim();
const solutionVersionOf = (value: string | undefined) => value?.match(/\(([\d.]+)\)\s*$/)?.[1];

const UNMANAGED_OWNERS = new Set(['active', 'default']);
const PLATFORM_OWNERS = /^(system|basic)$/i;
const MICROSOFT_OWNERS = /^(ms[a-z0-9]*_|msdyn|mscrm|msft|microsoft|powervirtualagents|basecustomcontrols|dataflowsolution|environmentvariables|connectionreference|appcommon|customcontrols)/i;
const MICROSOFT_COMPONENT = /^(msdyn|msdynce|mscrm|msft|mspp|msfp|adx)_/i;
// Component types the parser itemises from a ZIP; other types fall back to name matching across types.
const ITEMISED_TYPES = new Set([
  'table', 'column', 'relationship', 'global option set', 'form', 'view', 'process', 'app', 'web resource',
  'environment variable', 'connection reference', 'plugin assembly', 'plugin step', 'security role',
  'field security profile', 'email template', 'report', 'offline profile', 'custom api',
]);

interface ProviderIndex {
  ids: Set<string>;
  keys: Set<string>;
  shellTables: Set<string>;
}

function compareVersions(left: string, right: string): number {
  const a = left.split('.').map(Number);
  const b = right.split('.').map(Number);
  for (let index = 0; index < Math.max(a.length, b.length); index += 1) {
    const difference = (a[index] ?? 0) - (b[index] ?? 0);
    if (difference) return difference;
  }
  return 0;
}

export function describeRequired(dependency: MissingDependency): string {
  const name = /^[0-9a-f-]{36}$/i.test(dependencyKey(dependency.requiredName)) && dependency.requiredDisplayName
    ? dependency.requiredDisplayName : dependency.requiredName;
  const qualified = dependency.requiredParent ? `${dependency.requiredParent}.${name}`
    : dependency.requiredParentDisplayName ? `${name} (${dependency.requiredParentDisplayName})` : name;
  const label = dependency.requiredDisplayName && dependencyKey(dependency.requiredDisplayName) !== dependencyKey(name)
    && dependencyKey(dependency.requiredDisplayName) !== dependencyKey(qualified) ? ` "${dependency.requiredDisplayName}"` : '';
  return `${dependency.requiredType}: ${qualified}${label}`;
}

export function describeDependent(dependency: MissingDependency): string {
  const name = /^[0-9a-f-]{36}$/i.test(dependencyKey(dependency.dependentName)) && dependency.dependentDisplayName
    ? dependency.dependentDisplayName : dependency.dependentName;
  return `${dependency.dependentType}: ${name}`;
}

function buildIndex(solution: ParsedSolution, trackedTypes: Set<string>): ProviderIndex {
  const index: ProviderIndex = { ids: new Set(), keys: new Set(), shellTables: new Set() };
  const id = (value: string | undefined) => { if (value) index.ids.add(dependencyKey(value)); };
  const add = (type: string, name: string | undefined, parent = '') => {
    if (!name) return;
    trackedTypes.add(dependencyKey(type));
    index.keys.add(`${dependencyKey(type)}|${dependencyKey(parent)}|${dependencyKey(name)}`);
    index.keys.add(`*|${dependencyKey(parent)}|${dependencyKey(name)}`);
  };
  solution.ownedComponentIds?.forEach(id);
  solution.solutionComponents?.forEach((component) => {
    id(component.objectId);
    add(component.componentTypeName, component.name, component.entityLogicalName);
    if (component.componentType === 1 && component.rootComponentBehavior === 2) index.shellTables.add(dependencyKey(component.name));
  });
  solution.entities.forEach((entity) => {
    add('Table', entity.logicalName);
    add('Table', entity.name);
    entity.attributes.forEach((attribute) => add('Column', attribute.name, entity.logicalName));
    entity.relationships.forEach((relationship) => add('Relationship', relationship.name));
  });
  const fullTables = new Set(solution.entities.map((entity) => dependencyKey(entity.logicalName))
    .filter((name) => !(solution.solutionComponents ?? []).some((component) => component.componentType === 1
      && dependencyKey(component.name) === name && component.rootComponentBehavior === 2)));
  index.shellTables.forEach((table) => { if (fullTables.has(table)) index.shellTables.delete(table); });
  solution.optionSets.forEach((choice) => { id(choice.metadataId); add('Global Option Set', choice.name); });
  solution.forms.forEach((form) => [form.name, form.displayName].forEach((name) => add('Form', name, form.entityLogicalName)));
  solution.views.forEach((view) => [view.name, view.displayName].forEach((name) => add('View', name, view.entityLogicalName)));
  solution.processes.forEach((process) => [process.name, process.uniqueName, process.displayName].forEach((name) => add('Process', name)));
  solution.apps.forEach((app) => [app.name, app.uniqueName, app.displayName].forEach((name) => add('App', name)));
  solution.webResources.forEach((resource) => [resource.name, resource.schemaName].forEach((name) => add('Web Resource', name)));
  solution.environmentVariables.forEach((variable) => [variable.name, variable.schemaName].forEach((name) => add('Environment Variable', name)));
  solution.connectionReferences.forEach((reference) => [reference.name, reference.displayName].forEach((name) => add('Connection Reference', name)));
  solution.pluginAssemblies.forEach((assembly) => {
    [assembly.name, assembly.assemblyName].forEach((name) => add('Plugin Assembly', name));
    assembly.steps.forEach((step) => add('Plugin Step', step.name));
  });
  const named: Array<[string, Array<{ name: string; displayName?: string }>]> = [
    ['Security Role', solution.securityRoles], ['Field Security Profile', solution.fieldSecurityProfiles],
    ['Email Template', solution.emailTemplates], ['Report', solution.reports], ['Form', solution.dashboards],
    ['Offline Profile', solution.offlineProfiles], ['Custom API', solution.customApis], ['Copilot', solution.agents],
    ['AI Model', solution.aiModels], ['Desktop Flow', solution.desktopFlows], ['Dataflow', solution.dataflows],
  ];
  named.forEach(([type, items]) => items.forEach((item) => [item.name, item.displayName].forEach((name) => add(type, name))));
  return index;
}

export function analyzeDependencies(solutions: ParsedSolution[]): DependencyAnalysis {
  const identities = new Map(solutions.map((solution) => [
    `${dependencyKey(solution.metadata.uniqueName)}|${solution.metadata.version}|${solution.metadata.isManaged}`, solution,
  ]));
  const selected = [...identities.values()].sort((left, right) => left.metadata.uniqueName.localeCompare(right.metadata.uniqueName));
  const trackedTypes = new Set(ITEMISED_TYPES);
  const indexes = new Map(selected.map((solution) => [solution, buildIndex(solution, trackedTypes)]));
  const customPrefixes = new Set(selected.flatMap((solution) => [solution.metadata.publisherPrefix, solution.metadata.solutionPrefix])
    .filter((prefix): prefix is string => Boolean(prefix)).map(dependencyKey));
  const notes = new Set<string>();

  const solutionsNamed = (name: string) => {
    const unique = selected.filter((candidate) => dependencyKey(candidate.metadata.uniqueName) === dependencyKey(name));
    if (unique.length) return unique;
    const friendly = selected.filter((candidate) => dependencyKey(candidate.metadata.displayName) === dependencyKey(name));
    if (friendly.length === 1) notes.add(`Solution reference ${name} matched the display name of ${friendly[0].metadata.uniqueName}.`);
    return friendly.length === 1 ? friendly : [];
  };

  const indexFor = (solution: ParsedSolution) => indexes.get(solution) ?? buildIndex(solution, trackedTypes);
  const match = (solution: ParsedSolution, dependency: MissingDependency): 'full' | 'shell' | undefined => {
    const index = indexFor(solution);
    if ([dependency.requiredId, ...(dependency.requiredKeys ?? [])].some((value) => value && index.ids.has(dependencyKey(value)))) return 'full';
    const type = dependencyKey(dependency.requiredType);
    const parents = dependency.requiredParent ? [dependencyKey(dependency.requiredParent)]
      : dependency.requiredParentDisplayName ? selected.flatMap((candidate) => candidate.entities)
        .filter((entity) => dependencyKey(entity.displayName) === dependencyKey(dependency.requiredParentDisplayName))
        .map((entity) => dependencyKey(entity.logicalName)) : [''];
    const names = [dependency.requiredName, ...(dependency.requiredKeys ?? []), dependency.requiredDisplayName].map(dependencyKey).filter(Boolean);
    const typeKeys = trackedTypes.has(type) ? [type] : [type, '*'];
    const found = parents.some((parent) => names.some((name) => typeKeys.some((candidate) => index.keys.has(`${candidate}|${parent}|${name}`))));
    if (!found) return undefined;
    return type === 'table' && names.some((name) => index.shellTables.has(name)) ? 'shell' : 'full';
  };

  const hasCustomPrefix = (name: string) => name.includes('_') && customPrefixes.has(name.split('_')[0]);
  const standardTable = (name: string) => !!name && !name.includes('_');
  const isStandard = (dependency: MissingDependency) => {
    const name = dependencyKey(dependency.requiredName);
    const parent = dependencyKey(dependency.requiredParent);
    if (hasCustomPrefix(name) || hasCustomPrefix(parent)) return false;
    if (dependency.requiredType === 'Table') return standardTable(name);
    if (dependency.requiredType === 'Column') return standardTable(parent) && (!name.includes('_') || /^address\d_/.test(name));
    if (dependency.requiredType === 'Global Option Set') return standardTable(name);
    return false;
  };
  const isMicrosoftComponent = (dependency: MissingDependency) =>
    MICROSOFT_COMPONENT.test(dependencyKey(dependency.requiredName)) || MICROSOFT_COMPONENT.test(dependencyKey(dependency.requiredParent));

  const classify: DependencyAnalysis['classify'] = (consumer, dependency) => {
    const ownerName = solutionNameOf(dependency.requiredSolution);
    const ownerKey = dependencyKey(ownerName);
    const unmanagedOwner = !ownerKey || UNMANAGED_OWNERS.has(ownerKey);
    const owner = dependency.requiredPackages?.length
      ? dependency.requiredPackages.map((item) => `${item.appName}${item.version ? ` ${item.version}` : ''}`).join(', ')
      : dependency.requiredSolution;
    const consumerName = dependency.solutionUniqueName ?? consumer.metadata.uniqueName;
    if (!unmanagedOwner && ownerKey === dependencyKey(consumerName)) {
      return { status: 'included', providers: [], owner, reason: 'The manifest names this solution as the owner.' };
    }
    const selfMatch = match(consumer, dependency);
    if (selfMatch === 'full') return { status: 'included', providers: [], owner, reason: "Defined in this solution's ZIP." };
    const others = selected.filter((candidate) => candidate !== consumer
      && dependencyKey(candidate.metadata.uniqueName) !== dependencyKey(consumerName));
    const packageNames = new Set((dependency.requiredPackages ?? []).map((item) => dependencyKey(item.appName)));
    const named = unmanagedOwner ? [] : others.filter((candidate) => solutionsNamed(ownerName).includes(candidate)
      || packageNames.has(dependencyKey(candidate.metadata.uniqueName)));
    if (named.length) {
      const exported = solutionVersionOf(dependency.requiredSolution);
      named.forEach((provider) => {
        if (exported && /^\d+(\.\d+)*$/.test(provider.metadata.version) && compareVersions(provider.metadata.version, exported) < 0) {
          notes.add(`${consumerName} was exported against ${provider.metadata.uniqueName} ${exported}; the selected ZIP is ${provider.metadata.version}. Verify compatibility.`);
        }
      });
      const verified = named.some((provider) => match(provider, dependency));
      if (!verified && trackedTypes.has(dependencyKey(dependency.requiredType))) {
        notes.add(`${describeRequired(dependency)} (needed by ${consumerName}) is owned by selected solution ${named.map((provider) => provider.metadata.uniqueName).join(', ')}, but was not found in that ZIP. Verify the selected version contains it.`);
      }
      return { status: 'selected', providers: named, owner, reason: verified
        ? 'Owner declared in the manifest and found in its selected ZIP.'
        : 'Owner declared in the manifest is selected; the component was not itemised in its ZIP, so the owner declaration is used (see Verification Notes).' };
    }
    const found = others.filter((candidate) => match(candidate, dependency));
    const full = found.filter((candidate) => match(candidate, dependency) === 'full');
    if (found.length) {
      const providers = full.length ? full : found;
      return { status: 'selected', providers, owner, reason: providers.length > 1
        ? 'Found in several selected ZIPs; install after all of them.' : 'Found in a selected ZIP.' };
    }
    if (selfMatch === 'shell') return { status: 'included', providers: [], owner, reason: 'Included in this solution as a table shell.' };
    if (dependency.requiredPackages?.length) {
      return { status: 'environment', providers: [], owner, reason: `Delivered by managed application ${owner}. Install or verify it in the target environment before importing.` };
    }
    if (!unmanagedOwner) {
      if (PLATFORM_OWNERS.test(ownerName)) return { status: 'platform', providers: [], owner, reason: 'Standard Dataverse component owned by the platform.' };
      if (MICROSOFT_OWNERS.test(ownerName) || isMicrosoftComponent(dependency)) {
        return { status: 'environment', providers: [], owner, reason: `Supplied by Microsoft managed solution ${dependency.requiredSolution}. Verify it is installed in the target environment.` };
      }
      return { status: 'missing', providers: [], owner, reason: `Owned by managed solution ${dependency.requiredSolution}, which is not among the selected ZIPs. Install it in the target environment first, or add its ZIP and regenerate.` };
    }
    if (isStandard(dependency)) {
      return { status: 'platform', providers: [], owner, reason: ownerKey
        ? 'Standard Dataverse component (customised unmanaged in the source environment; those customisations are not carried by this solution).'
        : 'Standard Dataverse component present in every environment.' };
    }
    if (isMicrosoftComponent(dependency)) {
      return { status: 'environment', providers: [], owner, reason: 'Microsoft first-party component. Verify the owning Dynamics 365/Power Platform application is installed in the target environment.' };
    }
    if (ownerKey && ['Form', 'View', 'Chart'].includes(dependency.requiredType)
      && (standardTable(dependencyKey(dependency.requiredParent)) || (!dependency.requiredParent && !!dependency.requiredParentDisplayName))) {
      return { status: 'environment', providers: [], owner, reason: 'Unmanaged form/view on a standard table in the source environment. Verify it exists in the target environment, or add it to a solution.' };
    }
    return { status: 'missing', providers: [], owner, reason: ownerKey
      ? 'Unmanaged component in the source environment ("Active" layer) that is not in any selected ZIP. Add it to a solution, export, and include that ZIP.'
      : 'Not found in any selected ZIP and not recognised as a standard component. Add the solution that contains it.' };
  };

  const resolutions = selected.flatMap((solution) => collectMissingDependencies(solution).map((dependency) => ({
    consumer: dependency.solutionUniqueName ?? solution.metadata.uniqueName,
    consumerSolution: solution,
    dependency,
    ...classify(solution, dependency),
  })));
  return { solutions: selected, resolutions, get notes() { return [...notes].sort(); }, classify, solutionsNamed };
}
