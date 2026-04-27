import JSZip from 'jszip';
import { parseSolutionZip } from '../src/parser/solutionParser';
import { generateMarkdown } from '../src/generator/markdownGenerator';

async function buildSyntheticSolutionZip(): Promise<Buffer> {
  const zip = new JSZip();

  const solutionXml = `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <SolutionManifest>
    <UniqueName>ppmd_smoke</UniqueName>
    <Version>1.0.0.0</Version>
    <Managed>false</Managed>
    <LocalizedNames>
      <LocalizedName description="PPMD Smoke Solution" languagecode="1033" />
    </LocalizedNames>
    <Publisher>
      <UniqueName>ppmd</UniqueName>
      <LocalizedNames>
        <LocalizedName description="PPMD" languagecode="1033" />
      </LocalizedNames>
    </Publisher>
  </SolutionManifest>
</ImportExportXml>`;

  const customizationsXml = `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml>
  <Workflows>
    <Workflow>
      <Name>Order_Processor_Flow</Name>
      <Category>6</Category>
      <State>1</State>
      <PrimaryEntity>account</PrimaryEntity>
      <LocalizedNames>
        <LocalizedName description="Order Processor Flow" languagecode="1033" />
      </LocalizedNames>
    </Workflow>
  </Workflows>
  <connectionreferences>
    <connectionreference>
      <Name>cr_dataverse</Name>
      <connectionreferencedisplayname>Dataverse Main Connection</connectionreferencedisplayname>
      <connectorid>/providers/Microsoft.PowerApps/apis/shared_commondataserviceforapps</connectorid>
      <connectionid>/providers/Microsoft.PowerApps/connections/shared-commondataserviceforapps-123</connectionid>
    </connectionreference>
  </connectionreferences>
  <AppModules>
    <AppModule>
      <UniqueName>contoso_sales_hub</UniqueName>
      <Name>Contoso Sales Hub</Name>
      <ClientVersion>1.2.3.4</ClientVersion>
      <IsEnabled>true</IsEnabled>
      <LocalizedNames>
        <LocalizedName description="Contoso Sales Hub" languagecode="1033" />
      </LocalizedNames>
    </AppModule>
  </AppModules>
</ImportExportXml>`;

  const flowJson = {
    definition: {
      triggers: {
        When_a_row_is_added: {
          type: 'OpenApiConnection',
          inputs: {
            host: {
              apiId: '/providers/Microsoft.PowerApps/apis/shared_commondataserviceforapps',
            },
          },
        },
      },
      actions: {
        Scope_Main: {
          type: 'Scope',
          actions: {
            Condition_Check: {
              type: 'If',
              actions: {
                If_True_Action: {
                  type: 'OpenApiConnection',
                  inputs: {
                    host: {
                      apiId: '/providers/Microsoft.PowerApps/apis/shared_office365users',
                    },
                  },
                },
              },
              else: {
                actions: {
                  If_False_Action: {
                    type: 'OpenApiConnection',
                    inputs: {
                      host: {
                        apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
                      },
                    },
                  },
                },
              },
            },
            Switch_Path: {
              type: 'Switch',
              cases: {
                CaseA: {
                  actions: {
                    Case_A_Action: {
                      type: 'OpenApiConnection',
                      inputs: {
                        host: {
                          apiId: '/providers/Microsoft.PowerApps/apis/shared_teams',
                        },
                      },
                    },
                  },
                },
              },
              default: {
                actions: {
                  Default_Action: {
                    type: 'Compose',
                  },
                },
              },
            },
          },
        },
      },
    },
  };

  zip.file('solution.xml', solutionXml);
  zip.file('customizations.xml', customizationsXml);
  zip.file('Workflows/Order_Processor_Flow.json', JSON.stringify(flowJson, null, 2));

  return zip.generateAsync({ type: 'nodebuffer' });
}

function assertContains(markdown: string, token: string): void {
  if (!markdown.includes(token)) {
    throw new Error(`Missing expected token: ${token}`);
  }
}

async function run(): Promise<void> {
  const zipBytes = await buildSyntheticSolutionZip();
  const parsed = await parseSolutionZip(zipBytes as unknown as Blob);
  const markdown = generateMarkdown(parsed);

  const requiredTokens = [
    'Condition Check',
    'If True Action',
    'If False Action',
    'Case A Action',
    'Default Action',
    'Microsoft Dataverse',
    'Office 365 Users',
    'SharePoint',
    'Microsoft Teams',
    'Power Apps',
    'Unique Name',
    'contoso_sales_hub',
  ];

  requiredTokens.forEach((token) => assertContains(markdown, token));

  console.log('Smoke doc validation passed.');
  console.log(`Processes parsed: ${parsed.processes.length}`);
  console.log(`Connection refs parsed: ${parsed.connectionReferences.length}`);
  console.log(`Apps parsed: ${parsed.apps.length}`);
}

run().catch((err) => {
  console.error('Smoke doc validation failed.');
  console.error(err);
  process.exit(1);
});
