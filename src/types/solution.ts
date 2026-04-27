/**
 * @file solution.ts
 * @description Core TypeScript type definitions representing every component
 * that can appear inside a Power Platform solution ZIP file.  These types are
 * the backbone of the parser and the documentation generator.
 */

// ---------------------------------------------------------------------------
// Enumerations
// ---------------------------------------------------------------------------

/**
 * The category of a Power Platform solution component as stored in the
 * solution.xml <ComponentType> element.
 */
export enum ComponentType {
  // Data layer
  Entity = 'Entity',
  Attribute = 'Attribute',
  Relationship = 'Relationship',
  OptionSet = 'OptionSet',

  // Forms & Views
  Form = 'Form',
  SavedQuery = 'SavedQuery',          // View
  SystemForm = 'SystemForm',

  // Processes & automation
  Workflow = 'Workflow',              // Classic workflow / BPF / Power Automate
  PowerAutomateFlow = 'PowerAutomateFlow',

  // Apps
  AppModule = 'AppModule',           // Model-driven app
  CanvasApp = 'CanvasApp',           // Canvas app or Custom Page
  AIPlugin = 'AIPlugin',             // Code / AI apps

  // UI Resources
  WebResource = 'WebResource',

  // Security
  Role = 'Role',
  FieldSecurityProfile = 'FieldSecurityProfile',

  // Integration
  ConnectionReference = 'ConnectionReference',
  EnvironmentVariableDefinition = 'EnvironmentVariableDefinition',
  EnvironmentVariableValue = 'EnvironmentVariableValue',
  ServiceEndpoint = 'ServiceEndpoint',
  Connector = 'Connector',

  // Reporting
  Report = 'Report',
  Dashboard = 'Dashboard',

  // Code
  PluginAssembly = 'PluginAssembly',
  PluginStep = 'PluginStep',

  // Other
  SdkMessageProcessingStep = 'SdkMessageProcessingStep',
  Unknown = 'Unknown',
}

/**
 * The type of a Power Automate / classic workflow process.
 */
export enum ProcessCategory {
  Workflow = 'Workflow',
  Dialog = 'Dialog',
  BusinessRule = 'BusinessRule',
  Action = 'Action',
  BusinessProcessFlow = 'BusinessProcessFlow',
  CustomAction = 'CustomAction',
  PowerAutomateFlow = 'PowerAutomateFlow',
}

/**
 * The type of Power App.
 */
export enum AppType {
  ModelDriven = 'ModelDriven',
  Canvas = 'Canvas',
  CustomPage = 'CustomPage',
  CodeApp = 'CodeApp',
  AIPlugin = 'AIPlugin',
}

/**
 * The data type of an entity column/attribute.
 */
export enum AttributeType {
  String = 'String',
  Integer = 'Integer',
  Decimal = 'Decimal',
  Money = 'Money',
  Boolean = 'Boolean',
  DateTime = 'DateTime',
  Lookup = 'Lookup',
  OptionSet = 'OptionSet',
  MultiSelectOptionSet = 'MultiSelectOptionSet',
  Owner = 'Owner',
  UniqueIdentifier = 'UniqueIdentifier',
  Memo = 'Memo',
  Image = 'Image',
  File = 'File',
  Customer = 'Customer',
  PartyList = 'PartyList',
  Virtual = 'Virtual',
  BigInt = 'BigInt',
  ManagedProperty = 'ManagedProperty',
  Unknown = 'Unknown',
}

/**
 * The type of a web resource.
 */
export enum WebResourceType {
  HTML = 'HTML',
  CSS = 'CSS',
  JavaScript = 'JavaScript',
  TypeScript = 'TypeScript',
  XML = 'XML',
  PNG = 'PNG',
  JPG = 'JPG',
  GIF = 'GIF',
  XAP = 'XAP',
  XSL = 'XSL',
  ICO = 'ICO',
  SVG = 'SVG',
  Resx = 'Resx',
  Unknown = 'Unknown',
}

// ---------------------------------------------------------------------------
// Base types
// ---------------------------------------------------------------------------

/**
 * A named item with an optional description.
 */
export interface NamedItem {
  /** Unique name / schema name */
  name: string;
  /** Display / friendly name */
  displayName?: string;
  /** Description text */
  description?: string;
}

// ---------------------------------------------------------------------------
// Entity / Table layer
// ---------------------------------------------------------------------------

/**
 * An option within a choice (OptionSet) column.
 */
export interface OptionSetOption {
  /** Numeric integer value stored in Dataverse */
  value: number;
  /** Display label */
  label: string;
  /** Optional description */
  description?: string;
  /** Colour hint (hex, used in some customisations) */
  color?: string;
}

/**
 * A global or local OptionSet (Choice) definition.
 */
export interface OptionSetDefinition extends NamedItem {
  /** Whether the set is global (reusable across entities) */
  isGlobal: boolean;
  /** The individual options */
  options: OptionSetOption[];
}

/**
 * A column/attribute on a Dataverse table.
 */
export interface EntityAttribute extends NamedItem {
  /** Data type of the column */
  type: AttributeType;
  /** Whether the column is required */
  required: boolean;
  /** Whether the column is auditable */
  auditing?: boolean;
  /** Max length for text fields */
  maxLength?: number;
  /** Precision for numeric fields */
  precision?: number;
  /** For Lookup columns – the target entity logical name */
  lookupTarget?: string;
  /** For OptionSet columns – inline options when not referencing global set */
  options?: OptionSetOption[];
  /** For OptionSet columns – name of the global OptionSet if referenced */
  optionSetName?: string;
  /** Whether this is a custom attribute */
  isCustom: boolean;
  /** Whether this is the primary name attribute (used as the record label) */
  isPrimaryName?: boolean;
  /** Whether auditing is enabled for this column */
  isAuditEnabled?: boolean;
}

/**
 * Relationship between two Dataverse tables.
 */
export interface EntityRelationship {
  /** Schema name of the relationship */
  name: string;
  /** Relationship cardinality */
  type: 'OneToMany' | 'ManyToMany' | 'ManyToOne';
  /** Logical name of the related/referenced entity */
  referencedEntity: string;
  /** Logical name of the referencing entity */
  referencingEntity: string;
  /** Attribute on the referencing entity that holds the lookup/key */
  referencingAttribute?: string;
  /** Attribute on the referenced (parent) entity of the relationship */
  referencedAttribute?: string;
  /** Delete cascade behaviour */
  cascadeDelete?: string;
  /** Assign cascade behaviour */
  cascadeAssign?: string;
  /** Reparent cascade behaviour */
  cascadeReparent?: string;
  /** Human-readable description of the relationship */
  relationshipDescription?: string;
}

/**
 * A Dataverse table (entity) with all of its metadata.
 */
export interface EntityDefinition extends NamedItem {
  /** Logical name (e.g. account, new_widget) */
  logicalName: string;
  /** Object Type Code */
  objectTypeCode?: number;
  /** Whether the entity is custom */
  isCustom: boolean;
  /** Whether the entity is activity-enabled */
  isActivity?: boolean;
  /** Whether change tracking is enabled */
  changeTracking?: boolean;
  /** Columns / attributes defined in the solution */
  attributes: EntityAttribute[];
  /** Relationships defined in the solution */
  relationships: EntityRelationship[];
  /** Keys (alternate keys) */
  keys?: string[];
  /** Owner type: User or Organization */
  ownershipType?: 'User' | 'Organization' | 'None';
  /** OData entity set name (collection endpoint) */
  entitySetName?: string;
  /** Logical name of the primary name attribute */
  primaryAttributeName?: string;
}

// ---------------------------------------------------------------------------
// Forms & Views
// ---------------------------------------------------------------------------

/**
 * A field displayed on a form.
 */
export interface FormField {
  /** Attribute logical name */
  attributeName: string;
  /** Column label shown on the form */
  label?: string;
  /** Whether the field is required on the form */
  required?: boolean;
}

/**
 * A Dataverse form (main, quick view, quick create, etc.)
 */
export interface FormDefinition extends NamedItem {
  /** Entity the form belongs to */
  entityLogicalName: string;
  /** Form type: Main, QuickView, QuickCreate, Card, etc. */
  formType: string;
  /** Fields visible on the form */
  fields: FormField[];
}

/**
 * A Dataverse view (saved query).
 */
export interface ViewDefinition extends NamedItem {
  /** Entity the view belongs to */
  entityLogicalName: string;
  /** View type: Public, System, Quick Find, etc. */
  viewType: string;
  /** Columns shown in the view */
  columns: string[];
  /** FetchXML query behind the view */
  fetchXml?: string;
}

// ---------------------------------------------------------------------------
// Processes
// ---------------------------------------------------------------------------

/**
 * A single step/action within a workflow or flow.
 */
export interface ProcessStep {
  /** Step identifier */
  id: string;
  /** Human-readable name */
  name: string;
  /** Step type: SetAttributeValue, SendEmail, CallAction, etc. */
  stepType: string;
  /** Description or notes */
  description?: string;
  /** Child steps (for composite actions) */
  children?: ProcessStep[];
  /** Dataverse tables read or written by this step (Power Automate flows) */
  referencedEntities?: string[];
}

/**
 * A Power Platform process (classic workflow, business rule, BPF,
 * Power Automate flow, custom action, etc.)
 */
export interface ProcessDefinition extends NamedItem {
  /** Schema name */
  uniqueName: string;
  /** Category of process */
  category: ProcessCategory;
  /** Primary entity the process targets */
  primaryEntity?: string;
  /** All tables referenced by the process, where discoverable */
  relatedEntities?: string[];
  /**
   * Whether the process is activated.
   * Undefined means activation status is not present in the exported solution ZIP.
   */
  isActivated?: boolean;
  /** Trigger type: OnDemand, OnCreate, OnChange, etc. */
  triggerType?: string;
  /** Attributes that trigger the process (for OnChange) */
  triggerAttributes?: string[];
  /** Process steps extracted from the XML/JSON definition */
  steps: ProcessStep[];
  /** Run-as information */
  runAs?: string;
  /** Scope: User, BusinessUnit, etc. */
  scope?: string;
  /** Raw JSON definition for Power Automate flows */
  flowDefinition?: object;
  /** Description of trigger for Power Automate flows */
  flowTrigger?: string;
  /** List of connectors used in a Power Automate flow */
  flowConnectors?: string[];
  /** Connection reference logical/display names used in a Power Automate flow */
  flowConnectionReferences?: string[];
  /** Environment variable schema/display names used in a Power Automate flow */
  flowEnvironmentVariables?: string[];
}

// ---------------------------------------------------------------------------
// Apps
// ---------------------------------------------------------------------------

/**
 * A Power App (canvas, model-driven, custom page, code app).
 */
export interface AppDefinition extends NamedItem {
  /** App type */
  appType: AppType;
  /** Unique name */
  uniqueName: string;
  /** Entities (tables) the app surfaces */
  entities?: string[];
  /** Sitemap or navigation structure (model-driven) */
  sitemapAreas?: string[];
  /** Connectors used (canvas apps) */
  connectors?: string[];
  /** Whether the app is enabled */
  isEnabled?: boolean;
  /** Version of the app (from manifest) */
  version?: string;
}

// ---------------------------------------------------------------------------
// Web Resources
// ---------------------------------------------------------------------------

/**
 * A web resource (JavaScript, TypeScript, HTML, CSS, image, etc.)
 */
export interface WebResourceDefinition extends NamedItem {
  /** Schema name including publisher prefix */
  schemaName: string;
  /** Resource type */
  resourceType: WebResourceType;
  /** Whether this web resource is enabled for mobile clients */
  enabledForMobile?: boolean;
  /** Whether this web resource is available offline */
  availableOffline?: boolean;
  /** Raw content (decoded from base64 or plain text) */
  content?: string;
  /** Content length in bytes */
  contentLength?: number;
}

// ---------------------------------------------------------------------------
// Security
// ---------------------------------------------------------------------------

/**
 * Privilege on a security role.
 */
export interface RolePrivilege {
  /** Entity or system privilege name */
  privilegeName: string;
  /** Access level: 0=None, 1=Basic, 2=Local, 3=Deep, 4=Global */
  depth: number;
}

/**
 * A Dataverse security role.
 */
export interface SecurityRoleDefinition extends NamedItem {
  /** Role privileges */
  privileges: RolePrivilege[];
  /** Whether role is inherited by child business units */
  inherited?: boolean;
}

/**
 * Column-level security profile.
 */
export interface FieldSecurityProfileDefinition extends NamedItem {
  /** Allowed read/update/create permissions per field */
  permissions: Array<{
    attributeName: string;
    canRead: boolean;
    canUpdate: boolean;
    canCreate: boolean;
  }>;
}

// ---------------------------------------------------------------------------
// Integration
// ---------------------------------------------------------------------------

/**
 * A connection reference within a solution.
 */
export interface ConnectionReferenceDefinition extends NamedItem {
  /** Connector logical name / api name */
  connectorId: string;
  /** Connection ID if resolved */
  connectionId?: string;
  /** Display name of the connector */
  connectorDisplayName?: string;
}

/**
 * An environment variable definition.
 */
export interface EnvironmentVariableDefinition extends NamedItem {
  /** Data type: String, Number, Boolean, JSON, DataSource, Secret */
  type: string;
  /** Default value defined in the solution */
  defaultValue?: string;
  /** Whether a value was provided */
  hasCurrentValue: boolean;
  /** Current/override value (masked if secret) */
  currentValue?: string;
  /** Schema name */
  schemaName: string;
}

/**
 * An email template included in the solution.
 */
export interface EmailTemplateDefinition extends NamedItem {
  /** Subject line, where available */
  subject?: string;
  /** Related Dataverse table/logical entity */
  entityLogicalName?: string;
  /** Template type/category */
  templateType?: string;
  /** LCID language code */
  languageCode?: string;
}

// ---------------------------------------------------------------------------
// Reporting
// ---------------------------------------------------------------------------

/**
 * A Power Platform / SSRS Report.
 */
export interface ReportDefinition extends NamedItem {
  /** Report file name */
  fileName?: string;
  /** Related entities */
  relatedEntities?: string[];
  /** Report category */
  category?: string;
}

/**
 * A dashboard definition.
 */
export interface DashboardDefinition extends NamedItem {
  /** Entity context (or global) */
  entityLogicalName?: string;
  /** Dashboard type: Standard, FieldService, etc. */
  dashboardType?: string;
  /** Component IDs on the dashboard */
  components: string[];
}

// ---------------------------------------------------------------------------
// Plugins
// ---------------------------------------------------------------------------

/**
 * A registered plugin step.
 */
export interface PluginStepDefinition {
  /** Step unique name */
  name: string;
  /** SDK message (Create, Update, Delete, etc.) */
  message: string;
  /** Primary entity filter */
  primaryEntity?: string;
  /** Stage: PreValidation=10, PreOperation=20, PostOperation=40 */
  stage: number;
  /** Mode: Synchronous=0, Asynchronous=1 */
  mode: number;
  /** Class handling the step */
  pluginTypeName: string;
  /** Filtering attributes (comma-separated) */
  filteringAttributes?: string;
  /** Description */
  description?: string;
}

/**
 * A plugin assembly registered in the solution.
 */
export interface PluginAssemblyDefinition extends NamedItem {
  /** Assembly file name */
  assemblyName: string;
  /** Version string */
  version?: string;
  /** Culture */
  culture?: string;
  /** Public key token */
  publicKeyToken?: string;
  /** Isolation mode: None=1, Sandbox=2 */
  isolationMode: number;
  /** Source location: Database, Disk, GAC */
  sourceType?: string;
  /** Plugin types and their steps */
  steps: PluginStepDefinition[];
}

// ---------------------------------------------------------------------------
// Top-level parsed solution
// ---------------------------------------------------------------------------

/**
 * Metadata parsed from solution.xml at the root of the solution ZIP.
 */
export interface SolutionMetadata {
  /** Unique name of the solution */
  uniqueName: string;
  /** Friendly display name */
  displayName: string;
  /** Solution version string */
  version: string;
  /** Publisher display name */
  publisherName: string;
  /** Publisher unique name */
  publisherUniqueName?: string;
  /** Description */
  description?: string;
  /** Managed/Unmanaged flag */
  isManaged: boolean;
}

/**
 * The fully-parsed representation of a Power Platform solution ZIP.
 * This is the central data model passed to the Markdown generator.
 */
export interface ParsedSolution {
  /** Metadata from solution.xml */
  metadata: SolutionMetadata;
  /** Tables / entities */
  entities: EntityDefinition[];
  /** Global option sets */
  optionSets: OptionSetDefinition[];
  /** Forms */
  forms: FormDefinition[];
  /** Views */
  views: ViewDefinition[];
  /** Processes (classic workflows, BPFs, Power Automate, Business Rules) */
  processes: ProcessDefinition[];
  /** Power Apps */
  apps: AppDefinition[];
  /** Web Resources */
  webResources: WebResourceDefinition[];
  /** Security roles */
  securityRoles: SecurityRoleDefinition[];
  /** Column-level security profiles */
  fieldSecurityProfiles: FieldSecurityProfileDefinition[];
  /** Connection references */
  connectionReferences: ConnectionReferenceDefinition[];
  /** Environment variables */
  environmentVariables: EnvironmentVariableDefinition[];
  /** Email templates */
  emailTemplates: EmailTemplateDefinition[];
  /** Reports */
  reports: ReportDefinition[];
  /** Dashboards */
  dashboards: DashboardDefinition[];
  /** Plugin assemblies (with embedded steps) */
  pluginAssemblies: PluginAssemblyDefinition[];
  /** Any parse warnings or non-fatal errors */
  warnings: string[];
}
