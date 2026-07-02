import JSZip from 'jszip';

import { parseSolutionZip } from '../parser/solutionParser';
import { AppType, AttributeType } from '../types/solution';

async function buildTestZip(): Promise<Blob> {
  const zip = new JSZip();

  zip.file(
    'solution.xml',
    `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>contoso_test</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>contoso</UniqueName>
      <FriendlyName>Contoso</FriendlyName>
    </Publisher>
    <RootComponents>
      <RootComponent type="20" />
      <RootComponent type="36" />
    </RootComponents>
  </SolutionManifest>
</ImportExportXml>`,
  );

  zip.file('customizations.xml', '<ImportExportXml />');

  zip.file(
    'environmentvariabledefinitions/contoso_feature/environmentvariabledefinition.xml',
    `<?xml version="1.0" encoding="utf-8"?>
<environmentvariabledefinition>
  <environmentvariabledefinitionid>{11111111-1111-1111-1111-111111111111}</environmentvariabledefinitionid>
  <schemaname>contoso_FeatureFlag</schemaname>
  <displayname>Feature Flag</displayname>
  <type>String</type>
  <defaultvalue>default-enabled</defaultvalue>
</environmentvariabledefinition>`,
  );

  zip.file(
    'environmentvariablevalues/feature-flag.xml',
    `<?xml version="1.0" encoding="utf-8"?>
<environmentvariablevalue>
  <EnvironmentVariableDefinitionId>{11111111-1111-1111-1111-111111111111}</EnvironmentVariableDefinitionId>
  <value>environment-enabled</value>
</environmentvariablevalue>`,
  );

  const msapp = new JSZip();
  msapp.file(
    'Src/Home.fx.yaml',
    'Home As screen:\n  Fill: =RGBA(255, 255, 255, 1)\n  OnVisible: =Navigate(Settings, ScreenTransition.None)',
  );
  msapp.file(
    'Src/Settings.fx.yaml',
    'Settings As screen:\n  Fill: =RGBA(250, 250, 250, 1)',
  );

  const msappBytes = await msapp.generateAsync({ type: 'uint8array' });
  zip.file('CanvasApps/contoso_custompage.msapp', msappBytes);
  zip.file(
    'CanvasApps/contoso_custompage.json',
    JSON.stringify({
      name: 'contoso_custompage',
      displayName: 'Contoso Custom Page',
      appType: 'customPage',
      properties: { displayName: 'Contoso Custom Page' },
    }),
  );

  return zip.generateAsync({ type: 'blob' });
}

describe('parseSolutionZip regressions', () => {
  it('maps component types, captures canvas screens, and resolves environment variable values', async () => {
    const blob = await buildTestZip();
    const parsed = await parseSolutionZip(blob);

    const inventoryTypes = new Set(parsed.metadata.componentInventory.map((item) => item.componentType));
    expect(inventoryTypes.has('Role')).toBe(true);
    expect(inventoryTypes.has('EmailTemplate')).toBe(true);
    expect(Array.from(inventoryTypes).some((item) => item.includes('Unknown (20)') || item.includes('Unknown (36)'))).toBe(false);

    const customPage = parsed.apps.find((app) => app.uniqueName.toLowerCase() === 'contoso_custompage');
    expect(customPage).toBeTruthy();
    expect(customPage?.appType).toBe(AppType.CustomPage);
    expect(customPage?.canvasInsights?.screenCount).toBe(2);
    expect(customPage?.canvasInsights?.screenNames ?? []).toEqual(['Home', 'Settings']);

    const envVar = parsed.environmentVariables.find((item) => item.schemaName === 'contoso_FeatureFlag');
    expect(envVar).toBeTruthy();
    expect(envVar?.defaultValue).toBe('default-enabled');
    expect(envVar?.hasCurrentValue).toBe(true);
    expect(envVar?.currentValue).toBe('environment-enabled');
  });

  it('does not treat IsCustomizable alone as a custom table flag', async () => {
    const zip = new JSZip();
    zip.file(
      'solution.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>contoso_entity_customizable</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>contoso</UniqueName>
      <FriendlyName>Contoso</FriendlyName>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`,
    );

    zip.file(
      'customizations.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Entities>
    <Entity>
      <Name LocalizedName="Contact">contact</Name>
      <EntityInfo>
        <entity Name="contact" DisplayName="Contact">
          <IsCustomizable>
            <Value>true</Value>
          </IsCustomizable>
        </entity>
      </EntityInfo>
      <EntityRelationships />
    </Entity>
  </Entities>
</ImportExportXml>`,
    );

    const blob = await zip.generateAsync({ type: 'blob' });
    const parsed = await parseSolutionZip(blob);
    const entity = parsed.entities.find((item) => item.logicalName === 'contact');

    expect(entity).toBeTruthy();
    expect(entity?.isCustom).toBe(false);
  });

  it('extracts form field locations for body/header/footer controls', async () => {
    const zip = new JSZip();
    zip.file(
      'solution.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>contoso_forms</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>contoso</UniqueName>
      <FriendlyName>Contoso</FriendlyName>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`,
    );

    zip.file(
      'customizations.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Forms>
    <SystemForm>
      <Name>Account Main</Name>
      <ObjectTypeCode>account</ObjectTypeCode>
      <Type>Main</Type>
      <FormXml><![CDATA[
        <form>
          <tabs>
            <tab name="General">
              <columns>
                <column>
                  <sections>
                    <section name="Summary">
                      <rows>
                        <row>
                          <cell>
                            <control id="name" />
                          </cell>
                        </row>
                      </rows>
                    </section>
                  </sections>
                </column>
              </columns>
            </tab>
          </tabs>
          <header>
            <rows>
              <row>
                <cell>
                  <control id="ownerid" />
                </cell>
              </row>
            </rows>
          </header>
          <footer>
            <rows>
              <row>
                <cell>
                  <control id="statuscode" />
                </cell>
              </row>
            </rows>
          </footer>
        </form>
      ]]></FormXml>
    </SystemForm>
  </Forms>
</ImportExportXml>`,
    );

    const blob = await zip.generateAsync({ type: 'blob' });
    const parsed = await parseSolutionZip(blob);

    const form = parsed.forms.find((item) => item.name === 'Account Main');
    expect(form).toBeTruthy();

    const nameField = form?.fields.find((field) => field.attributeName === 'name');
    expect(nameField).toBeTruthy();
    expect(nameField?.location).toBe('body');
    expect(nameField?.tabName).toBe('General');
    expect(nameField?.sectionName).toBe('Summary');

    const ownerField = form?.fields.find((field) => field.attributeName === 'ownerid');
    expect(ownerField?.location).toBe('header');

    const statusField = form?.fields.find((field) => field.attributeName === 'statuscode');
    expect(statusField?.location).toBe('footer');
  });

  it('extracts agents, AI models, desktop flows, and enriched attribute metadata', async () => {
    const zip = new JSZip();
    zip.file(
      'solution.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>contoso_phase4</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>contoso</UniqueName>
      <FriendlyName>Contoso</FriendlyName>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`,
    );

    zip.file(
      'customizations.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Entities>
    <Entity>
      <Name LocalizedName="Account">account</Name>
      <EntityInfo>
        <entity Name="account" DisplayName="Account" />
      </EntityInfo>
      <attributes>
        <attribute Name="new_budget" Type="Decimal">
          <DisplayName>
            <LocalizedLabels>
              <LocalizedLabel languagecode="1033" description="Budget" />
            </LocalizedLabels>
          </DisplayName>
          <RequiredLevel Value="None" />
          <MinValue>1</MinValue>
          <MaxValue>1000</MaxValue>
          <Precision>2</Precision>
          <DefaultValue>100</DefaultValue>
          <Format>Currency</Format>
        </attribute>
        <attribute Name="new_regarding" Type="Lookup">
          <RequiredLevel Value="None" />
          <Targets>
            <Target>account</Target>
            <Target>contact</Target>
          </Targets>
        </attribute>
      </attributes>
      <EntityRelationships />
    </Entity>
  </Entities>
</ImportExportXml>`,
    );

    zip.file(
      'Agents/sales-assistant.json',
      JSON.stringify({
        displayName: 'Sales Assistant',
        agentType: 'TaskAgent',
        language: 'en-US',
        trigger: 'Teams',
        connectors: ['Dataverse'],
      }),
    );

    zip.file(
      'AIModels/case-classifier.json',
      JSON.stringify({
        displayName: 'Case Classifier',
        modelType: 'Classification',
        provider: 'Azure OpenAI',
        version: '2026.1',
        endpoint: 'contoso-ai',
      }),
    );

    zip.file(
      'DesktopFlows/invoice-reconciliation.json',
      JSON.stringify({
        displayName: 'Invoice Reconciliation',
        folder: 'Finance',
        enabled: true,
        steps: [{ name: 'Launch app' }, { name: 'Read data' }],
        connectors: ['SAP'],
      }),
    );

    const blob = await zip.generateAsync({ type: 'blob' });
    const parsed = await parseSolutionZip(blob);

    expect(parsed.agents.length).toBe(1);
    expect(parsed.agents[0].displayName).toBe('Sales Assistant');
    expect(parsed.aiModels.length).toBe(1);
    expect(parsed.aiModels[0].provider).toBe('Azure OpenAI');
    expect(parsed.desktopFlows.length).toBe(1);
    expect(parsed.desktopFlows[0].stepCount).toBe(2);

    const entity = parsed.entities.find((item) => item.logicalName === 'account');
    expect(entity).toBeTruthy();
  });

  it('parses custom and advanced-find flags from metadata key variants', async () => {
    const zip = new JSZip();
    zip.file(
      'solution.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>contoso_attrflags</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>contoso</UniqueName>
      <FriendlyName>Contoso</FriendlyName>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`,
    );

    zip.file(
      'customizations.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Entities>
    <Entity>
      <Name LocalizedName="Account">account</Name>
      <EntityInfo>
        <entity Name="account" DisplayName="Account">
          <attributes>
            <attribute Name="new_customflag" Type="String">
              <IsCustomAttribute>
                <Value>1</Value>
              </IsCustomAttribute>
              <ValidForAdvancedFind>1</ValidForAdvancedFind>
            </attribute>
            <attribute Name="new_notadvanced" Type="String">
              <IsCustomAttribute>
                <Value>true</Value>
              </IsCustomAttribute>
              <IsValidForAdvancedFind>
                <Value>false</Value>
              </IsValidForAdvancedFind>
            </attribute>
            <attribute Name="accountnumber" Type="String">
              <IsCustomizable>
                <Value>true</Value>
              </IsCustomizable>
              <IsValidForAdvancedFind>true</IsValidForAdvancedFind>
            </attribute>
          </attributes>
        </entity>
      </EntityInfo>
      <EntityRelationships />
    </Entity>
  </Entities>
</ImportExportXml>`,
    );

    const blob = await zip.generateAsync({ type: 'blob' });
    const parsed = await parseSolutionZip(blob);
    const entity = parsed.entities.find((item) => item.logicalName === 'account');

    expect(entity).toBeTruthy();
    const customFlag = entity?.attributes.find((attr) => attr.name === 'new_customflag');
    const notAdvanced = entity?.attributes.find((attr) => attr.name === 'new_notadvanced');
    const accountNumber = entity?.attributes.find((attr) => attr.name === 'accountnumber');

    expect(customFlag?.isCustom).toBe(true);
    expect(customFlag?.isValidForAdvancedFind).toBe(true);
    expect(notAdvanced?.isCustom).toBe(true);
    expect(notAdvanced?.isValidForAdvancedFind).toBe(false);
    expect(accountNumber?.isCustom).toBe(false);
    expect(accountNumber?.isValidForAdvancedFind).toBe(true);
    expect(customFlag?.metadataSources?.isCustom).toContain('IsCustomAttribute');
    expect(customFlag?.metadataSources?.isValidForAdvancedFind).toContain('ValidForAdvancedFind');
    expect(accountNumber?.metadataSources?.isValidForAdvancedFind).toContain('IsValidForAdvancedFind');
  });

  it('parses lowercase metadata key variants without marking default columns as custom', async () => {
    const zip = new JSZip();
    zip.file(
      'solution.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>contoso_attrflags_lowercase</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>contoso</UniqueName>
      <FriendlyName>Contoso</FriendlyName>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`,
    );

    zip.file(
      'customizations.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Entities>
    <Entity>
      <Name LocalizedName="Account">account</Name>
      <EntityInfo>
        <entity Name="account" DisplayName="Account">
          <attributes>
            <attribute Name="new_lowercasecustom" Type="String">
              <iscustomattribute>
                <value>true</value>
              </iscustomattribute>
              <validforadvancedfind>
                <value>true</value>
              </validforadvancedfind>
            </attribute>
            <attribute Name="name" Type="String">
              <iscustomizable>
                <value>true</value>
              </iscustomizable>
              <isvalidforadvancedfind>
                <value>true</value>
              </isvalidforadvancedfind>
            </attribute>
          </attributes>
        </entity>
      </EntityInfo>
      <EntityRelationships />
    </Entity>
  </Entities>
</ImportExportXml>`,
    );

    const blob = await zip.generateAsync({ type: 'blob' });
    const parsed = await parseSolutionZip(blob);
    const entity = parsed.entities.find((item) => item.logicalName === 'account');

    expect(entity).toBeTruthy();
    const lowercaseCustom = entity?.attributes.find((attr) => attr.name === 'new_lowercasecustom');
    const systemName = entity?.attributes.find((attr) => attr.name === 'name');

    expect(lowercaseCustom?.isCustom).toBe(true);
    expect(lowercaseCustom?.isValidForAdvancedFind).toBe(true);
    expect(systemName?.isCustom).toBe(false);
    expect(systemName?.isValidForAdvancedFind).toBe(true);
    expect(lowercaseCustom?.metadataSources?.isCustom).toContain('iscustomattribute');
    expect(systemName?.metadataSources?.isValidForAdvancedFind).toContain('isvalidforadvancedfind');
  });

  it('parses IsCustomField and DisplayMask-based advanced-find metadata', async () => {
    const zip = new JSZip();
    zip.file(
      'solution.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>contoso_attrflags_displaymask</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>contoso</UniqueName>
      <FriendlyName>Contoso</FriendlyName>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`,
    );

    zip.file(
      'customizations.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Entities>
    <Entity>
      <Name LocalizedName="Account">account</Name>
      <EntityInfo>
        <entity Name="account" DisplayName="Account">
          <attributes>
            <attribute Name="new_fromdisplaymask" Type="String">
              <IsCustomField>1</IsCustomField>
              <DisplayMask>ValidForAdvancedFind|ValidForForm|ValidForGrid</DisplayMask>
            </attribute>
            <attribute Name="createdon" Type="DateTime">
              <IsCustomField>0</IsCustomField>
              <DisplayMask>ValidForForm|ValidForGrid</DisplayMask>
            </attribute>
          </attributes>
        </entity>
      </EntityInfo>
      <EntityRelationships />
    </Entity>
  </Entities>
</ImportExportXml>`,
    );

    const blob = await zip.generateAsync({ type: 'blob' });
    const parsed = await parseSolutionZip(blob);
    const entity = parsed.entities.find((item) => item.logicalName === 'account');

    expect(entity).toBeTruthy();
    const fromDisplayMask = entity?.attributes.find((attr) => attr.name === 'new_fromdisplaymask');
    const createdOn = entity?.attributes.find((attr) => attr.name === 'createdon');

    expect(fromDisplayMask?.isCustom).toBe(true);
    expect(fromDisplayMask?.isValidForAdvancedFind).toBe(true);
    expect(createdOn?.isCustom).toBe(false);
    expect(createdOn?.isValidForAdvancedFind).toBeUndefined();
    expect(fromDisplayMask?.metadataSources?.isCustom).toContain('IsCustomField');
    expect(fromDisplayMask?.metadataSources?.isValidForAdvancedFind).toContain('DisplayMask');
  });

  it('parses lookup type from AttributeTypeName and infers custom flag from publisher prefix fallback', async () => {
    const zip = new JSZip();
    zip.file(
      'solution.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>acoe_attr_lookup</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>acoe</UniqueName>
      <FriendlyName>ACOE</FriendlyName>
      <CustomizationPrefix>acoe</CustomizationPrefix>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`,
    );

    zip.file(
      'customizations.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Entities>
    <Entity>
      <Name LocalizedName="Case">incident</Name>
      <EntityInfo>
        <entity Name="incident" DisplayName="Case">
          <attributes>
            <attribute Name="acoe_lotrapplicationid">
              <AttributeTypeName>
                <Value>LookupType</Value>
              </AttributeTypeName>
              <RequiredLevel>
                <Value>SystemRequired</Value>
              </RequiredLevel>
            </attribute>
            <attribute Name="title" Type="String">
              <RequiredLevel>
                <Value>None</Value>
              </RequiredLevel>
            </attribute>
          </attributes>
        </entity>
      </EntityInfo>
      <EntityRelationships />
    </Entity>
  </Entities>
</ImportExportXml>`,
    );

    const blob = await zip.generateAsync({ type: 'blob' });
    const parsed = await parseSolutionZip(blob);
    const entity = parsed.entities.find((item) => item.logicalName === 'incident');

    expect(entity).toBeTruthy();
    const lookupAttr = entity?.attributes.find((attr) => attr.name === 'acoe_lotrapplicationid');
    const titleAttr = entity?.attributes.find((attr) => attr.name === 'title');

    expect(lookupAttr?.type).toBe(AttributeType.Lookup);
    expect(lookupAttr?.required).toBe(true);
    expect(lookupAttr?.isCustom).toBe(true);

    expect(titleAttr?.type).toBe(AttributeType.String);
    expect(titleAttr?.isCustom).toBe(false);
  });

  it('parses lookup type from AttributeTypeDisplayName and Targets when Type is missing', async () => {
    const zip = new JSZip();
    zip.file(
      'solution.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>acoe_attr_lookup_displayname</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>acoe</UniqueName>
      <FriendlyName>ACOE</FriendlyName>
      <CustomizationPrefix>acoe</CustomizationPrefix>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`,
    );

    zip.file(
      'customizations.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Entities>
    <Entity>
      <Name LocalizedName="Case">incident</Name>
      <EntityInfo>
        <entity Name="incident" DisplayName="Case">
          <attributes>
            <attribute Name="acoe_lorapplicationid">
              <AttributeTypeDisplayName>
                <Value>Lookup</Value>
              </AttributeTypeDisplayName>
              <Targets>
                <Target>acoe_lorapplication</Target>
              </Targets>
            </attribute>
          </attributes>
        </entity>
      </EntityInfo>
      <EntityRelationships />
    </Entity>
  </Entities>
</ImportExportXml>`,
    );

    const blob = await zip.generateAsync({ type: 'blob' });
    const parsed = await parseSolutionZip(blob);
    const entity = parsed.entities.find((item) => item.logicalName === 'incident');
    const lookupAttr = entity?.attributes.find((attr) => attr.name === 'acoe_lorapplicationid');

    expect(lookupAttr?.type).toBe(AttributeType.Lookup);
    expect(lookupAttr?.lookupTarget).toBe('acoe_lorapplication');
    expect(lookupAttr?.isCustom).toBe(true);
  });
});
