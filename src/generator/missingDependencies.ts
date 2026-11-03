import type { MissingDependency, ParsedSolution } from '../types/solution';

export const MISSING_DEPENDENCIES_NOTE = 'Import prerequisites declared in the solution ZIP MissingDependencies manifest, classified against this solution, the other selected ZIPs, the standard Dataverse platform and Microsoft-supplied solutions. Only items marked Missing are absent from all of these.';

const normalize = (value: string) => value.trim().replace(/^\{|\}$/g, '').toLowerCase();

export function dependencyComponentType(type: number | string): string {
  const dependencyLabels: Record<number, string> = {
    1: 'Table', 2: 'Column', 3: 'Relationship', 9: 'Global Option Set', 10: 'Relationship',
    20: 'Security Role', 24: 'Form', 26: 'View', 29: 'Process', 31: 'Report',
    36: 'Email Template', 59: 'Chart', 60: 'Form', 61: 'Web Resource', 62: 'Site Map',
    66: 'Custom Control', 68: 'Custom Control Default Config', 70: 'Field Security Profile', 80: 'App',
    90: 'Plugin Type', 91: 'Plugin Assembly', 92: 'Plugin Step',
    161: 'Offline Profile', 300: 'App', 371: 'Connector', 372: 'Connector',
    380: 'Environment Variable', 381: 'Environment Variable Value', 10039: 'Connection Reference',
  };
  const textLabels: Record<string, string> = {
    connectionreference: 'Connection Reference', environmentvariabledefinition: 'Environment Variable',
    environmentvariablevalue: 'Environment Variable Value', settingdefinition: 'Setting Definition', appsetting: 'App Setting',
  };
  const raw = String(type).trim();
  if (/^\d+$/.test(raw)) return dependencyLabels[Number(raw)] ?? `Component ${raw}`;
  return textLabels[raw.toLowerCase()] ?? (raw || 'Component');
}

export function missingDependenciesNotice(solution: ParsedSolution): string | undefined {
  if (solution.missingDependenciesError) return solution.missingDependenciesError;
  if (solution.missingDependencies === undefined) {
    return 'Import dependencies have not been collected for this solution. Regenerate from the solution ZIP.';
  }
  return undefined;
}

export function collectMissingDependencies(solution: ParsedSolution): MissingDependency[] {
  const sourceSolutions = new Set((solution.dependencySolutionUniqueNames ?? [solution.metadata.uniqueName]).map(normalize));
  const ownersByItem = new Map<string, string[]>();
  const registerOwners = (type: string, names: (string | undefined)[], owners: string[]) => {
    names.filter((name): name is string => Boolean(name)).forEach((name) => {
      const key = `${type}|${normalize(name)}`;
      ownersByItem.set(key, [...new Set([...(ownersByItem.get(key) ?? []), ...owners])]);
    });
  };
  solution.forms.forEach((form) => registerOwners('Form', [form.name, form.displayName], [form.entityLogicalName]));
  solution.views.forEach((view) => registerOwners('View', [view.name, view.displayName], [view.entityLogicalName]));
  solution.apps.forEach((app) => registerOwners('App', [app.name, app.displayName, app.uniqueName], app.entities ?? []));
  solution.processes.forEach((process) => registerOwners('Process', [process.name, process.displayName, process.uniqueName],
    [process.primaryEntity ?? '', ...(process.relatedEntities ?? [])].filter(Boolean)));
  solution.pluginAssemblies.forEach((assembly) => assembly.steps.forEach((step) =>
    registerOwners('Plugin Step', [step.name, `${assembly.assemblyName} :: ${step.name}`], step.primaryEntity ? [step.primaryEntity] : [])));

  const unique = new Map<string, MissingDependency>();
  (solution.missingDependencies ?? []).filter((dependency) => dependency.requiredName.trim()
    && (!dependency.solutionUniqueName || sourceSolutions.has(normalize(dependency.solutionUniqueName)))).forEach((dependency) => {
    const solutionUniqueName = dependency.solutionUniqueName ?? solution.metadata.uniqueName;
    const key = [
      solutionUniqueName, dependency.dependentType, dependency.dependentId || dependency.dependentName,
      dependency.requiredType, dependency.requiredParent ?? '', dependency.requiredId || dependency.requiredName, dependency.requiredSolution ?? '',
    ].map(normalize).join('|');
    const existing = unique.get(key);
    unique.set(key, {
      ...dependency,
      solutionUniqueName,
      tableLogicalNames: [...new Set([
        ...(existing?.tableLogicalNames ?? []), ...dependency.tableLogicalNames,
        ...(ownersByItem.get(`${dependency.dependentType}|${normalize(dependency.dependentName)}`) ?? []),
      ].map(normalize).filter(Boolean))].sort(),
    });
  });
  return [...unique.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([, dependency]) => dependency);
}

export function dependenciesForTable(dependencies: MissingDependency[], logicalName: string): MissingDependency[] {
  return dependencies.filter((dependency) => dependency.tableLogicalNames.some((name) => normalize(name) === normalize(logicalName)));
}
