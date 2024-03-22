/*---------------------------------------------------------------------------------------------
*  Copyright (c) Microsoft Corporation. All rights reserved.
*  Licensed under the MIT License. See License.txt in the project root for license information.
*--------------------------------------------------------------------------------------------*/

/* eslint-disable @typescript-eslint/no-unsafe-assignment */

import * as fs from "fs-extra";
import { find } from "lodash";
import * as os from "os";
import * as path from "path";
import * as tmp from "tmp";
import * as vscode from "vscode";
import { Terminal } from "vscode";

import { ext } from "../extensionVariables";
import { SampleUrlInfo } from '../sample';
import { getCodeToCloudCommand } from "../subCommand/codeToCloudSlashCommand";
import {
  CREATE_SAMPLE_COMMAND_ID,
  createCommand,
  getCreateCommand,
} from "../subCommand/createSlashCommand";
import {
  getAgentHelpCommand,
  helpCommandName,
} from "../subCommand/helpSlashCommand";
import {
  DefaultNextStep,
  EXECUTE_COMMAND_ID,
  OPENURL_COMMAND_ID,
  executeCommand,
  getNextStepCommand,
  openUrlCommand,
} from "../subCommand/nextStep/command";
import { getTestCommand } from "../subCommand/testCommand";
import { buildFileTree, getSampleFileInfo, modifyFile } from "../util";
import { agentDescription, agentName, maxFollowUps, wxpAgentDescription, wxpAgentName, } from "./agentConsts";
import {
  LanguageModelID,
  getResponseAsStringCopilotInteraction
} from "./copilotInteractions";
import { SlashCommandHandlerResult, SlashCommandsOwner } from "./slashCommands";

export const CREATE_WXP_PROJECT_COMMAND_ID = 'teamsfx.createWxpProject';

export interface ITeamsChatAgentResult extends vscode.ChatResult {
  metadata: {
    slashCommand?: string;
    sampleIds?: string[];
  };
}

export type CommandVariables = {
  languageModelID?: LanguageModelID;
  chatMessageHistory?: vscode.LanguageModelChatMessage[];
};

export type AgentRequest = {
  slashCommand?: string;
  userPrompt: string;
  variables: readonly vscode.ChatResolvedVariable[];

  context: vscode.ChatContext;
  response: vscode.ChatExtendedResponseStream;
  token: vscode.CancellationToken;

  commandVariables?: CommandVariables;
};

export interface IAgentRequestHandler {
  handleRequestOrPrompt(
    request: AgentRequest
  ): Promise<SlashCommandHandlerResult>;
  getFollowUpForLastHandledSlashCommand(
    result: vscode.ChatResult,
    context: vscode.ChatContext,
    token: vscode.CancellationToken
  ): vscode.ChatFollowup[] | undefined;
}

/**
* Owns slash commands that are knowingly exposed to the user.
*/
const agentSlashCommandsOwner = new SlashCommandsOwner(
  {
    noInput: helpCommandName,
    default: defaultHandler,
  },
  { disableIntentDetection: true }
);
agentSlashCommandsOwner.addInvokeableSlashCommands(
  new Map([
    getCreateCommand(),
    getNextStepCommand(),
    getAgentHelpCommand(agentSlashCommandsOwner),
    getTestCommand(),
    getCodeToCloudCommand(),
  ])
);

export function registerChatAgent() {
  try {
    const participant = vscode.chat.createChatParticipant(agentName, handler);
    participant.description = agentDescription;
    participant.iconPath = vscode.Uri.joinPath(
      ext.context.extensionUri,
      "resources",
      "teams.png"
    );
    participant.followupProvider = { provideFollowups: followUpProvider };
    const wxpParticipant = vscode.chat.createChatParticipant(wxpAgentName, handler);
    wxpParticipant.iconPath = vscode.Uri.joinPath(
      ext.context.extensionUri,
      "resources",
      "M365.png"
    );
    wxpParticipant.description = wxpAgentDescription;
    wxpParticipant.followupProvider = { provideFollowups: followUpProvider };
    registerVSCodeCommands(participant, wxpParticipant);
  } catch (e) {
    console.log(e);
  }
}

async function handler(
  request: vscode.ChatRequest,
  context: vscode.ChatContext,
  stream: vscode.ChatResponseStream,
  token: vscode.CancellationToken
): Promise<vscode.ChatResult | undefined> {
  const agentRequest: AgentRequest = {
    slashCommand: request.command,
    userPrompt: request.prompt,
    variables: request.variables,
    context: context,
    response: stream,
    token: token,
  };
  let handleResult: SlashCommandHandlerResult | undefined;

  const handlers = [agentSlashCommandsOwner];
  for (const handler of handlers) {
    handleResult = await handler.handleRequestOrPrompt(agentRequest);
    if (handleResult !== undefined) {
      break;
    }
  }

  if (handleResult !== undefined) {
    handleResult.followUp = handleResult.followUp?.slice(0, maxFollowUps);
    return handleResult.chatAgentResult;
  } else {
    return undefined;
  }
}

function followUpProvider(
  result: ITeamsChatAgentResult,
  context: vscode.ChatContext,
  token: vscode.CancellationToken
): vscode.ProviderResult<vscode.ChatFollowup[]> {
  const providers = [agentSlashCommandsOwner];

  let followUp: vscode.ChatFollowup[] | undefined;
  for (const provider of providers) {
    followUp = provider.getFollowUpForLastHandledSlashCommand(
      result,
      context,
      token
    );
    if (followUp !== undefined) {
      break;
    }
  }
  followUp = followUp ?? [];
  if (followUp.length === 0) {
    followUp.push(DefaultNextStep);
  }
  return followUp;
}

function getCommands(
  _context: vscode.ChatContext,
  _token: vscode.CancellationToken
): vscode.ProviderResult<vscode.ChatCommand[]> {
  return agentSlashCommandsOwner.getSlashCommands().map(([name, config]) => ({
    name: name,
    description: config.shortDescription,
  }));
}

async function defaultHandler(
  request: AgentRequest
): Promise<SlashCommandHandlerResult> {
  let host = "";
  let codeMathToBeInserted = "";
  const srcRoot = os.homedir();
  const tmpDir = os.tmpdir();
  const defaultTargetFolder = srcRoot ? path.join(srcRoot, "Office-Add-in") : '';
  const tsfilePath = vscode.Uri.file(path.join(defaultTargetFolder, "src", "taskpane", "taskpane.ts"));
  const tsfilePathStr = path.join(defaultTargetFolder, "src", "taskpane", "taskpane.ts");
  const htmlfilePathStr = path.join(defaultTargetFolder, "src", "taskpane", "taskpane.html");
  const lastResponse = getLastResponse(request);
  const lastRequest = getLastRequest(request);
  const tmpRequestPath = path.join(tmpDir, 'tmpRequest.txt');
  const tmpCodePath = path.join(tmpDir, 'tmpCode.txt');
  // const tmpSummaryPath = path.join(tmpDir, 'tmpSummary.txt');
  // const tmpFolderPath = path.join(tmpDir, 'tmpFolder');
  // const intention = await analyzeIntention(request);
  const languageModelID: LanguageModelID = "copilot-gpt-3.5-turbo";
  const chatMessageHistory: vscode.LanguageModelChatMessage[] = [];
  const NextStepCreateDone: vscode.ChatFollowup = {
    prompt: "Create the project in the current workspace.",
    command: "",
    label: vscode.l10n.t("Create the project in the current workspace."),
  };
  const NextStepPublish: vscode.ChatFollowup = {
    prompt: "How can I distribute the add-in to more users?",
    command: "",
    label: vscode.l10n.t("How can I distribute the add-in to more users?"),
  };
  // if (lastRequest !== "" && lastResponse !== "") {
  //   chatMessageHistory.push(
  //     new vscode.LanguageModelUserMessage(lastRequest),
  //     new vscode.LanguageModelAssistantMessage(lastResponse)
  //   );
  // }
  request.commandVariables = { languageModelID, chatMessageHistory };
  const plannerPrompt = `
  I want you act as an expert in Office JavaScript add-in development area. All user asks related to Word, Excel or PowerPoint should be handled using Office JavaScript API Follow the <Instructions>.

  <Instructions>
  - You must categorize the user task into 2 types.
    1. General tasks. The general tasks are high-level and can be broken down into several steps.
      For example: "How to import data from a database to Excel?"
    2. Specific tasks. The specific tasks are low-level and can be finished in one-step action or a JavaScript method. Or it can be a sub-task of a general task.
      For example: "How to insert a table in Word?"
  - For general tasks, you should first introduce Office JavaScript add-in to the user, then tell the user how to finish the task step by step. The steps have the following structure:
    1. Tell user how to set up Office JavaScript add-in dev environment in VS Code.
    2. Generate code snippets of the method following <CodeStructure> for users to show how to finish the user task.
    3. Guide user to replace the existing run method.
    4. Guide user to update <HTMLElement> in the html file to include the new method name.
    5. Guide update the existing \`document.getElementById("[idName]").onclick = [functionName];\` properly.
    6. Tell user how to run npm install in the terminal and press F5 to debug the add-in.
  - For specific tasks, you should generate code snippets of the method following <CodeStructure> for users to show how to finish the user task.
  - You should use your knowledge in Office JavaScript add-in development area to help the user when necessary.
  </Instructions>

  <CodeStructure>
  - There must be one and only one main method in one code snippet. The main method must strictly follow the structure <CodeTemplate>.
  - The main method must have a meaningful [functionName], a correct [hostName] of Word, Excel or Powerpoint, and runnable [Code] to address the user's ask.
  - The main method should not have any passed in parameters. The necessary parameters should be defined inside the method.
  - The main method for each object should contain loading properties, get and set properties and some method calls. All the properties, method calls should be existing on this object or related with it.
  - Except for the main method, you can have other helper methods if necessary. All helper methods must be properly called in the main method.
  - No more code should be generated except for the methods.
  </CodeStructure>

  <CodeTemplate>
  \`\`\`javascript
  export async function [functionName]() {
    try {
      await [hostName]].run(async (context) => {
        [Code]
      })
    } catch (error) {
      console.error(error);
    }
  }
  \`\`\`
  </CodeTemplate>

  <HTMLElement>
    <div role="button" id="[idName]" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
        <span class="ms-Button-label">[buttonName]</span>
    </div>
  </HTMLElement>
  `;

  let fewShotSample = "";
  const fewShotSampleRange = `
  <ExcelSample>
  \`\`\`javascript
  sheet.getCell(0, 0).values = [[0]]; // Assign a 1*1 array to a single cell.
  sheet.getRange(\`A1:B3\`).values = [['Date', 'Close Price'], ['2024-01-01', 100], ['2024-01-02', 110]]; // Assign a 3*2 array to a 3*2 range.
  \`\`\`
</ExcelSample>`;
  const fewShotSampleFull = `
<ExcelSample>
    User: fetch stock data and import into Excel.
    GitHub Copilot:
    \`\`\`javascript
    //1. fetch data
      const symbol = 'MSFT';
      const key = 'YOUR_API';
      const response = await fetch(https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=$\{symbol\}&apikey=$\{key\});
  const data = await response.json();

  //2. parse data
  // Parse the JSON response and extract the necessary data
  // Replace the placeholders with your parsing logic
  const stockData = data['Time Series (Daily)'];
  const dates = Object.keys(stockData);
  const closePrices = Object.values(stockData).map(entry => entry['4. close']);
  let result = dates.map((item, index) => {
    return [item, closePrices[index]];
  });

  // 3. Import parsed data into Excel
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(\`A1: B$\{ result.length \}\`);
      range.values = result;
    \`\`\`
  </ExcelSample>
  `;

  if (request.userPrompt.includes('alpha')) {
    fewShotSample = fewShotSampleFull;
  }
  else {
    fewShotSample = fewShotSampleRange;
  }

  const stepByStepPrompt = `
  I want you act as an expert in Office JavaScript add-in development area. All user asks related to Word, Excel or PowerPoint should be handled using Office JavaScript API Follow the <Instructions>.

  <Instructions>
  - First, you should give a positive reply that  you understand the user's ask and tell user the task can be done by building an Office JavaScript add-in.
  - Second, you should summarize 2-3 features users need to implement to finish their task. Each featrue should be described in only a few words.
  - Then, generate a code sample following <CodeStructure> for users to show how to finish the user task. if an API key is needed, remember to notify the user.
  - Next, explain the code sample you just generated.
  - At the end of your response, you should ask user 'To run the code, you need to create an add-in project. Do you want to create a project in the current workspace?'.
  </Instructions>

  <CodeStructure>
  There must be one and only one main method in one code snippet. The main method must strictly follow the structure <CodeTemplate>.
  - The main method must have a meaningful [functionName], a correct [hostName] of Word, Excel or Powerpoint, and runnable [Code] to address the user's ask.
  - The main method should not have any passed in parameters. The necessary parameters should be defined inside the method.
  - The main method for each object should contain loading properties, get and set properties and some method calls. All the properties, method calls should be existing on this object or related with it.
  - All variable declarations MUST be in the body of the method.
  - When using REST API, you should use fetch.
  - Don't include any \`npm install\` command in your response.
  - When using Excel JavaScript API to set a [value] to range, you need to first clearly figure out the dimension of [value]. And then must make sure the dimension of the range must align with the dimension [value]. Take <ExcelSample> as an example.
  - Except for the main method, you can have other helper methods if necessary. All helper methods must be properly called in the main method.
  - No more code should be generated except for the methods.
  - The returned method should be well-implemented without any placeholder comments or fake functions.
  </CodeStructure>

  <CodeTemplate>
  \`\`\`javascript
  export async function [functionName]() {
    try {
      await [hostName]].run(async (context) => {
        [Code]
      })
    } catch (error) {
      console.error(error);
    }
  }
  \`\`\`
  </CodeTemplate>

  ${fewShotSample}
  `;

  const intentionPrompt = `
  Categorize the user intention into one of the 6 intentions below:
  1. "Ask for step-by-step guidance"
    For example:
    "I want to import stock data into Excel and do analysis. Tell me what to do."
    "I want to import NBA data into Excel and do analysis."
  2. "Show sample code"
    For example:
    "Format the table."
    "Generate a line chart."
  3. "Create a new project"
    For example:
    "Create the project in the current workspace."
  4. "Publish add-in"
    For example:
    "How can I distribute the add-in to more users?"
  5. "Fix the code"
    For example:
    "Fix the error"
    "I have an error:"
  6. "Others"
  Return the string of the intention only.
  `;
  // const codeTemplate = `

  // `

  let intentionResponse = await getResponseAsStringCopilotInteraction(intentionPrompt, request) ?? '';
  // request.response.markdown(intentionResponse);
  if (intentionResponse.includes("Ask for step-by-step guidance") || request.userPrompt.toLowerCase().includes("i want to")) {
    fs.unlink(tmpCodePath, (err) => {
      if (err) {
        console.log('Error deleting file:', err);
      } else {
        console.log('File deleted successfully');
      }
    });

    fs.unlink(tmpRequestPath, (err) => {
      if (err) {
        console.log('Error deleting file:', err);
      } else {
        console.log('File deleted successfully');
      }
    });

    // fs.unlink(tmpSummaryPath, (err) => {
    //   if (err) {
    //     console.log('Error deleting file:', err);
    //   } else {
    //     console.log('File deleted successfully');
    //   }
    // });
    // const summaryPrompt = `Summarize the user's purpose with less than 10 words without a subject.`;
    // let summaryResponse = await getResponseAsStringCopilotInteraction(summaryPrompt, request) ?? '';
    // await writeTextFile(tmpSummaryPath, summaryResponse);

    //request.commandVariables = { languageModelID: "copilot-gpt-4" };

    let response = await getResponseAsStringCopilotInteraction(stepByStepPrompt, request) ?? '';
    request.response.markdown(response);
    await writeTextFile(tmpRequestPath, request.userPrompt);
    let code = "";
    if (response !== "") {
      const regex = /```javascript([\s\S]*?)```/g;
      const matches = [...response.matchAll(regex)];
      code = matches.map((match) => match[1]).join('\n');
    }
    await writeTextFile(tmpCodePath, code);
    return { chatAgentResult: { metadata: { slashCommand: "create" } }, followUp: [NextStepCreateDone] };
  } else if (intentionResponse.includes("Show sample code")
    || (request.userPrompt.toLowerCase().includes('format')
      || request.userPrompt.toLowerCase().includes('chart')
      || request.userPrompt.toLowerCase().includes('alpha'))) {
    let lastCode = await readTextFile(tmpCodePath);
    const tsFileExist = await fileExists(tsfilePath);
    if (tsFileExist) {
      const generateCodePrompt = `
      I want you to generate Office JavaScript code following <Instructions> to resolve the user's ask.

      <Instructions>
      - If the ${lastCode} is not empty, you must generate code based on ${lastCode} and follow <CodeStructure>. If the ${lastCode} is empty, you should generate a new code snippet following <CodeStructure>.
      - Explain the code sample you just generated.
      </Instructions>

      <CodeStructure>
      - There must be one and only one main method in one code snippet. The main method must strictly follow the structure <CodeTemplate>.
      - The main method must have a meaningful [functionName], a correct [hostName] of Word, Excel or Powerpoint, and runnable [Code] to address the user's ask.
      - The main method should not have any passed in parameters. The necessary parameters should be defined inside the method.
      - The main method for each object should contain loading properties, get and set properties and some method calls. All the properties, method calls should be existing on this object or related with it.
      - All variable declarations MUST be in the body of the method.
      - When using REST API, you should use fetch. And Generate the code to fetch stock data from alphavantage.
      - Don't include any \`npm install\` command in your response.
      - When using Excel JavaScript API to generate some code. Take <ExcelSample> as an example.
      - Except for the main method, you can have other helper methods if necessary. All helper methods must be properly called in the main method.
      - No more code should be generated except for the methods.
      </CodeStructure>

      <CodeTemplate>
      \`\`\`javascript
      export async function [functionName]() {
        try {
          await [hostName]].run(async (context) => {
            [Code]
          })
        } catch (error) {
          console.error(error);
        }
      }
      \`\`\`
      </CodeTemplate>

      <ExcelSample>
      \`\`\`javascript
      // range is an existing variable of the type Excel.Range
      const chart = sheet.charts.add(Excel.ChartType.line, range, Excel.ChartSeriesBy.auto);
      chart.title.text = 'Stock Trend';
      chart.title.format.font.bold = true;
      \`\`\`
      </ExcelSample>
      ` ;
      const stepByStepRequest = await readTextFile(tmpRequestPath);
      let codeResponse = "";

      request.userPrompt = stepByStepRequest.split('.')[0] + '. ' + request.userPrompt;
      codeResponse = await getResponseAsStringCopilotInteraction(generateCodePrompt, request) ?? '';
      request.response.markdown(codeResponse);

      let code = "";
      if (codeResponse !== "") {
        const regex = /```javascript([\s\S]*?)```/g;
        const matches = [...codeResponse.matchAll(regex)];
        code = matches.map((match) => match[1]).join('\n');
      }
      if (code.length > lastCode.length) {
        await writeTextFile(tmpCodePath, code);
      }
      const inspirePrompt1 =
        `
      As an Office JavaScript Add-in expert, give an inspiration to the user what's the next step they can do in LESS than 10 words based on the user's request.
      - If the data request is format the table, suggest the user to generate a chart.
      - If the user request is generate a chart, suggest the user to add data labels to the chart.
      `
      const inspiration1 = await giveInspirationWithLLM(request, inspirePrompt1);
      const NextStep1: vscode.ChatFollowup = {
        prompt: inspiration1,
        command: "",
        label: vscode.l10n.t(inspiration1),
      };
      const inspirePrompt2 =
        `
      As an Office JavaScript Add-in expert, give an inspiration to the user what's the next step they can do in LESS than 10 words based on the user's request.
      - If the data request is format the table, suggest the user to format the data in another style.
      - If the user request is generate a chart, suggest the user to format the chart.
      `
      const inspiration2 = await giveInspirationWithLLM(request, inspirePrompt2);
      const NextStep2: vscode.ChatFollowup = {
        prompt: inspiration2,
        command: "",
        label: vscode.l10n.t(inspiration2),
      };
      return { chatAgentResult: { metadata: { slashCommand: "create" } }, followUp: [NextStep1, NextStep2, NextStepPublish] };
    } else {
      const generateCodePrompt = `
      I want you to generate Office JavaScript code following <Instructions> to resolve the user's ask.

      <Instructions>
      - If the ${lastCode} is not empty, you should generate code based on ${lastCode} and follow <CodeStructure>. If the ${lastCode} is empty, you should generate a new code snippet following <CodeStructure>.
      - Explain the code sample you just generated.
      - At the end of your response, you should ask user 'To run the code, you need to create an add-in project. Do you want to create a project in the current workspace?'.
      </Instructions>

      <CodeStructure>
      - There must be one and only one main method in one code snippet. The main method must strictly follow the structure <CodeTemplate>.
      - The main method must have a meaningful [functionName], a correct [hostName] of Word, Excel or Powerpoint, and runnable [Code] to address the user's ask.
      - The main method should not have any passed in parameters. The necessary parameters should be defined inside the method.
      - The main method for each object should contain loading properties, get and set properties and some method calls. All the properties, method calls should be existing on this object or related with it.
      - All variable declarations MUST be in the body of the method.
      - When using REST API, you should use fetch.
      - Don't include any \`npm install\` command in your response.
      - When using Excel JavaScript API to set the cell value, you should notice the dimension of the cell must be aligned with the dimension input array. Thus, you should figure out the dimension of the array first, and get the range of the cells. Take <ExcelSample> as an example.
      - Except for the main method, you can have other helper methods if necessary. All helper methods must be properly called in the main method.
      - No more code should be generated except for the methods.
      </CodeStructure>

      <CodeTemplate>
      \`\`\`javascript
      export async function [functionName]() {
        try {
          await [hostName]].run(async (context) => {
            [Code]
          })
        } catch (error) {
          console.error(error);
        }
      }
      \`\`\`
      </CodeTemplate>

      <ExcelSample>
      User: fetch stock data and import into Excel.
      GitHub Copilot:
      \`\`\`javascript
      //1. fetch data
      const symbol = 'MSFT';
      const key = 'YOUR_API';
      const response = await fetch(https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=$\{symbol\}&apikey=$\{key\});
      const data = await response.json();

      //2. parse data
      // Parse the JSON response and extract the necessary data
      // Replace the placeholders with your parsing logic
      const stockData = data['Time Series (Daily)'];
      const dates = Object.keys(stockData);
      const closePrices = Object.values(stockData).map(entry => entry['4. close']);
      let result = dates.map((item, index) => {
        return [item, closePrices[index]];
      });

      // 3. Import parsed data into Excel
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(\`A1: B$\{ result.length \}\`);
      range.values = result;
      \`\`\`
    </ExcelSample>
    `;

      const stepByStepRequest = await readTextFile(tmpRequestPath);
      let codeResponse = "";

      request.userPrompt = stepByStepRequest.split('.')[0] + '. ' + request.userPrompt;
      codeResponse = await getResponseAsStringCopilotInteraction(generateCodePrompt, request) ?? '';
      request.response.markdown(codeResponse);

      let code = "";
      if (codeResponse !== "") {
        const regex = /```javascript([\s\S]*?)```/g;
        const matches = [...codeResponse.matchAll(regex)];
        code = matches.map((match) => match[1]).join('\n');
      }
      if (code.length > lastCode.length) {
        await writeTextFile(tmpCodePath, code);
      }
      return { chatAgentResult: { metadata: { slashCommand: "create" } }, followUp: [NextStepCreateDone] };
    }
  } else if (intentionResponse.includes("Create a new project") || (request.userPrompt.toLowerCase().includes('y') && lastResponse.includes('create a project'))) {
    // if (vscode.workspace.workspaceFolders !== undefined && vscode.workspace.workspaceFolders.length > 0) {
    // const isFileExist = await fileExists(vscode.Uri.file(tmpFolderPath));
    const lastCode = await readTextFile(tmpCodePath);
    // if (!) {
    if (lastCode.includes('Excel')) {
      host = 'Excel';
    } else if (lastCode.includes('Word')) {
      host = 'Word';
    } else if (lastCode.includes('PowerPoint')) {
      host = 'PowerPoint';
    }
    codeMathToBeInserted = correctEnumSpelling(lastCode);

    const wxpSampleURLInfo: SampleUrlInfo = {
      owner: "GavinGu07",
      repository: "Office-Add-in-Templates",
      ref: "main",
      dir: host
    };
    const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(wxpSampleURLInfo, 2);
    const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
    const nodes = await buildFileTree(fileUrlPrefix, samplePaths, tempFolder, wxpSampleURLInfo.dir, 2, 20);

    const folder = path.join(tempFolder, wxpSampleURLInfo.dir);

    // fs.writeFile(tmpFolderPath, folder, (err) => {
    //   if (err) {
    //     console.log('Error writing file:', err);
    //   } else {
    //     console.log('File written successfully');
    //   }
    // });
    // const summary = await readTextFile(tmpSummaryPath);
    await modifyFile(folder, codeMathToBeInserted);
    request.response.markdown(`The ${host} add-in project has been created at ${defaultTargetFolder}.\n\n`);
    request.response.markdown(`The key files are:\n\n`);
    request.response.markdown(`1. **manifest.xml**: This is the manifest file for the Office Add-in. It defines the settings and capabilities of the add-in.\n\n`);
    request.response.markdown(`2. **package.json**: This is the configuration file for npm. It lists the dependencies and scripts for the project.\n\n`);
    request.response.markdown(`3. **src/ directory**: This directory contains the source code for the add-in.\n\n`)
    request.response.markdown(`\n\n To run the project, you need to first install all the packages needed:\n\n`);
    request.response.button({
      command: CREATE_WXP_PROJECT_COMMAND_ID,
      arguments: [folder, defaultTargetFolder],
      title: vscode.l10n.t('Create add-in project and install dependency')
    });
    // request.response.markdown(`\n\n Here is the tree structure of the add-in project.`);
    // request.response.filetree(nodes, vscode.Uri.file(path.join(tempFolder, wxpSampleURLInfo.dir)));
    // return { chatAgentResult: { slashCommand: "create" }, followUp: [NextStepCreateDone] };
    return { chatAgentResult: { metadata: { slashCommand: 'create' } }, followUp: [NextStepPublish] };
    // }
    // else {
    //   console.log('File exists');
    // const tmpFolder = await readTextFile(tmpFolderPath);
    // await fs.copy(tmpFolder, defaultTargetFolder);
    // fs.unlink(tmpFolderPath, (err) => {
    //   if (err) {
    //     console.log('Error deleting file:', err);
    //   } else {
    //     console.log('File deleted successfully');
    //   }
    // });
    // //   request.response.markdown(`The Office add-in project has been created at ${defaultTargetFolder}.`);
    // //   // const introduceProjectPrompt = `You should introduce the current workspace files`;
    // //   // request.userPrompt = '@workspace introduce the current workspace';
    // //   // let response = await getResponseAsStringCopilotInteraction(introduceProjectPrompt, request) ?? '';
    // //   // request.response.markdown(response);
    // //   // request.response.markdown(`\n\n To run the project, you need to first install all the packages needed:\n\n`);
    // //   // request.response.markdown(`\`\`\`bash\nnpm install\n\`\`\`\n`);
    // //   // request.response.markdown(`Then you can run the add-in project by hitting \`F5\` or running the following command:\n\n`);
    // //   // request.response.markdown(`\`\`\`bash\nnpm run start\n\`\`\`\n`);
    // //   request.response.button({
    // //     command: LAUNCH_TTK,
    // //     arguments: [],
    // //     title: vscode.l10n.t('Switch to Teams Toolkit Extension')
    // //   });

    // return { chatAgentResult: { slashCommand: "create" }, followUp: [NextStepPublish] };
    //   const fileExistingPrompt = `The current workspace already has an Office add-in project. You should guide user to create in a new workspace.`;
    //   let response = await getResponseAsStringCopilotInteraction(fileExistingPrompt, request) ?? '';
    //   request.response.markdown(response);
    //   return { chatAgentResult: { slashCommand: "" }, followUp: [] };
    // }
    // }
  } else if (intentionResponse.includes("Publish add-in")) {
    const publishAddInPrompt =
      `
    I want you to provide all documentations and steps to publish the Office add-in to the store and marketplace.
    `
    let response = await getResponseAsStringCopilotInteraction(publishAddInPrompt, request) ?? '';
    request.response.markdown(response);
    return { chatAgentResult: { metadata: { slashCommand: '' } }, followUp: [] };
  } else if (intentionResponse.includes("Fix the code")) {
    const fixCodePrompt =
      `
      The user has asked for help to fix the code. You should provide the user with the correct code to fix the issue using your knowledge in Office JavaScript APIs and Office Add-ins.
      `
    let response = await getResponseAsStringCopilotInteraction(fixCodePrompt, request) ?? '';
    request.response.markdown(response);
    return { chatAgentResult: { metadata: { slashCommand: 'create' } }, followUp: [NextStepPublish] };

  } else {
    const consultantPrompt =
      `
    You are an expert in Office JavaScript Add-in. Your job is to help the user learn about how they can use Office Add-in and Office JavaScript APIs to solve a problem or accomplish a task. Do not suggest using any other tools other than what has been previously mentioned. Assume the the user is only interested in using Office Add-in. Finally, do not overwhelm the user with too much information. Keep responses short and sweet.
    `
    let response = await getResponseAsStringCopilotInteraction(consultantPrompt, request) ?? '';
    request.response.markdown(response);
    return { chatAgentResult: { metadata: { slashCommand: '' } }, followUp: [] };
  }

  // let response = await getResponseAsStringCopilotInteraction(plannerPrompt, request) ?? '';
  // request.response.markdown(response);

  return { chatAgentResult: { metadata: { slashCommand: '' } }, followUp: [] };
  // console.log("defaultTargetFolder: " + defaultTargetFolder);
  // if (intention.includes('generate code')) {
  //   const objectJson = '{"Annotation":"Represents an annotation attached to a paragraph.","AnnotationCollection":"Contains a collection of Annotation objects.","Body":"Represents the body of a document or a section.","Border":"Represents the Border object for text, a paragraph, or a table.","BorderCollection":"Represents the collection of border styles.","CheckboxContentControl":"The data specific to content controls of type CheckBox.","Comment":"Represents a comment in the document.","CommentCollection":"Contains a collection of Comment objects.","CommentContentRange":"Specifies the comment\'s content range.","CommentReply":"Represents a comment reply in the document.","CommentReplyCollection":"Contains a collection of CommentReply objects. Represents all comment replies in one comment thread.","ContentControl":"Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, and checkbox content controls are supported.","ContentControlCollection":"Contains a collection of ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text and plain text content controls are supported.","CritiqueAnnotation":"Represents an annotation wrapper around critique displayed in the document.","CustomProperty":"Represents a custom property.","CustomPropertyCollection":"Contains the collection of CustomProperty objects.","CustomXmlPart":"Represents a custom XML part.","CustomXmlPartCollection":"Contains the collection of CustomXmlPart objects.","CustomXmlPartScopedCollection":"Contains the collection of CustomXmlPart objects with a specific namespace.","Document":"The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.","DocumentCreated":"The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.","DocumentProperties":"Represents document properties.","Field":"Represents a field.","FieldCollection":"Contains a collection of Field objects.","Font":"Represents a font.","InlinePicture":"Represents an inline picture.","InlinePictureCollection":"Contains a collection of InlinePicture objects.","List":"Contains a collection of Paragraph objects.","ListCollection":"Contains a collection of List objects.","ListItem":"Represents the paragraph list item format.","ListLevel":"Represents a list level.","ListLevelCollection":"Contains a collection of ListLevel objects.","ListTemplate":"Represents a ListTemplate.","NoteItem":"Represents a footnote or endnote.","NoteItemCollection":"Contains a collection of NoteItem objects.","Paragraph":"Represents a single paragraph in a selection, range, content control, or document body.","ParagraphCollection":"Contains a collection of Paragraph objects.","ParagraphFormat":"Represents a style of paragraph in a document.","Range":"Represents a contiguous area in a document.","RangeCollection":"Contains a collection of Range objects.","SearchOptions":"Specifies the options to be included in a search operation. To learn more about how to use search options in the Word JavaScript APIs, read Use search options to find text in your Word add-in.","Section":"Represents a section in a Word document.","SectionCollection":"Contains the collection of the document\'s Section objects.","Setting":"Represents a setting of the add -in.","SettingCollection":"Contains the collection of Setting objects.","Shading":"Represents the shading object.","Style":"Represents a style in a Word document.","StyleCollection":"Contains a collection of Style objects.","Table":"Represents a table in a Word document.","TableBorder":"Specifies the border style.","TableCell":"Represents a table cell in a Word document.","TableCellCollection":"Contains the collection of the document\'s TableCell objects.","TableCollection":"Contains the collection of the document\'s Table objects.","TableRow":"Represents a row in a Word document.","TableRowCollection":"Contains the collection of the document\'s TableRow objects.","TableStyle":"Represents the TableStyle object.","TrackedChange":"Represents a tracked change in a Word document.","TrackedChangeCollection":"Contains a collection of TrackedChange."}';
  //   const parsedObjectDescription = JSON.parse(objectJson);

  //   const generateProjectPrompt = `
  //       # Role
  //       I want you act as an expert in Office JavaScript add-in development area.You are also an advisor for Office add-in developers.

  //       # Instructions
  //       - Given the Office JavaScript add-in developer's request, please follow below to help determine the information about generating an JavaScript add-in project.
  //       - You should interpret the intention of developer's request as an ask to generate an Office JavaScript add-in project. And polish user input into some sentences if necessary.
  //       - You should go through the following steps silently, and only reply to user with a JSON result in each step. Do not explain why for your answer.

  //       - Suggest an platform for the add-in project.There are 3 options: Word, Excel, PowerPoint.If you can't determine, just say All.
  //       - You should base on your understanding of developer intent and the capabilities of Word, Excel, PowerPoint to make the suggestion.
  //       - Remember it as "PLATFORM".

  //       - Suggest an add-in type.You have 3 options: taskpane, content, custom function. You should notice Word doesn't have content type, and only Excel has custom function type. Remember it as "TYPE".

  //       - You should then base on the "PLATFORM" information and add-in developer asks to suggest one or a set of specific Office JavaScript API objects that are related.
  //       - You should analyze the API objects typical user cases or capabilities of their related UI features to suggest the most relevant ones.
  //       - The suggested API objects should not be too general such as "Document", "Workbook", "Presentation".
  //       - The suggested API objects should be from the list inside "API objects list".
  //       - The "API objects list" is a JSON object with a list of Office JavaScript API objects and their descriptions. The "API obejcts list" is as follows: ${JSON.stringify(parsedObjectDescription)}
  //       - You should give 3 most relevant objects. Remember it as "APISET".

  //       - Provide some detailed summary about why you make the suggestions in above steps. Remember it as "SUMMARY".
  //       ` ;

  //   const addinPlatfromTypeAPIResponse = await getResponseAsStringCopilotInteraction(generateProjectPrompt, request);
  //   if (addinPlatfromTypeAPIResponse) {
  //     const responseJson = parseCopilotResponseMaybeWithStrJson(addinPlatfromTypeAPIResponse);
  //     const apiObjectsStr = Array.isArray(responseJson.APISET) ? responseJson.APISET.map((api: string) => `${api}`).join(", ") : '';

  //     const generateCodePrompt = `
  //     # Role
  //     I want you act as an expert in Office JavaScript add-in development area.You are also an advisor for Office add-in developers.

  //     # Instructions
  //     - You should help generate a code snippet including Office JavaScript API calls based on user request.
  //     - The generated method must start with 'export async function' keyword.
  //     - The generated method should contain a meaningful function name and a runnable code snippet with its own context.
  //     - The generated method should have a try catch block to handle the exception.
  //     - Each generated method should contain Word.run, Excel.run or PowerPoint.run logic.
  //     - Each generated method should not have any passed in parameters. The necessary parameters should be defined inside the method.
  //     - The generated method for each object should contain loading properties, get and set properties and some method calls. All the properties, method calls should be existing on this object or related with it.
  //     - Remember to strictly reference the "API list" to generate the code. The "API list" is as follows: ${getApiListStringByObject(apiObjectsStr.split(', '))}.
  //     - If the userPrompt includes add or insert keywords, your generated code should contain insert or add method calls.
  //     `;

  //     let codeMath = "";

  //     const userRequestBackup = request.userPrompt;
  //     request.userPrompt = ` Please generate one method for each ${apiObjectsStr} ${responseJson.PLATFORM} JavaScript API object.`;
  //     host = `${responseJson.PLATFORM}`;
  //     while (codeMath === "") {
  //       const generatedCodeResponse = await getResponseAsStringCopilotInteraction(generateCodePrompt, request);
  //       if (generatedCodeResponse) {
  //         //const regex = new RegExp(`${quoteChar}(.*?)${quoteChar}`, 'g');
  //         const regex = /```javascript([\s\S]*?)```/g;
  //         const matches = [...generatedCodeResponse.matchAll(regex)];
  //         codeMath = matches.map((match) => match[1]).join('\n');

  //         console.log(codeMath);
  //       }
  //     }

  //     request.userPrompt = userRequestBackup;
  //     let codeMath2 = "";
  //     let generatedCodeResponse2: string | undefined = '';
  //     console.log(codeMath2);
  //     while (codeMath2 === "") {
  //       generatedCodeResponse2 = await getResponseAsStringCopilotInteraction(generateCodePrompt, request);
  //       if (generatedCodeResponse2) {
  //         //const regex = new RegExp(`${quoteChar}(.*?)${quoteChar}`, 'g');
  //         const regex = /```javascript([\s\S]*?)```/g;
  //         const matches = [...generatedCodeResponse2.matchAll(regex)];
  //         codeMath2 = matches.map((match) => match[1]).join('\n');
  //         console.log(codeMath2);
  //       }
  //     }
  //     codeMathToBeInserted = correctEnumSpelling(codeMath2);
  //     request.response.markdown(`${generatedCodeResponse2}`);
  //     request.response.markdown(`\n\nDo you want to try the code snippet in an Office add-in project?`);
  //   }
  //   const NextStepCreate: vscode.ChatFollowup = {
  //     prompt: "Create a new Office add-in including the above code snippet",
  //     command: "",
  //     label: vscode.l10n.t("Try the snippet in an Office add-in project"),
  //   };
  //   return { chatAgentResult: { slashCommand: "create" }, followUp: [NextStepCreate] };
  // }
  // else if (intention.includes('fix code')) {
  //   const activeTextEditor = vscode.window.activeTextEditor;
  //   if (activeTextEditor) {
  //     let uri = activeTextEditor.document.uri;
  //     let diagnostics = vscode.languages.getDiagnostics(uri);
  //     let errorDiagnostics = diagnostics.filter(diagnostic => diagnostic.severity === vscode.DiagnosticSeverity.Error);
  //     let diagnostic = errorDiagnostics[0];
  //     if (diagnostic) {
  //       let errorMessage = diagnostic.message;
  //       let errorCode = activeTextEditor.document.lineAt(diagnostic.range.start.line).text;
  //       await fixErrorCode(errorMessage, errorCode, request);
  //     }
  //   }
  //   return { chatAgentResult: { slashCommand: '' }, followUp: [] };
  // }
  // else if (lastResponse.includes("```javascript")) {
  //   // const lastTimeResponse: vscode.ChatRequestTurn | vscode.ChatResponseTurn | undefined = request.context.history.find(item => item instanceof vscode.ChatResponseTurn);
  //   // let response;
  //   // if (lastTimeResponse instanceof vscode.ChatResponseTurn) {
  //   //   response = lastTimeResponse.response;
  //   // }


  //   if (lastResponse.includes('Excel')) {
  //     host = 'Excel';
  //   } else if (lastResponse.includes('Word')) {
  //     host = 'Word';
  //   } else if (lastResponse.includes('PowerPoint')) {
  //     host = 'PowerPoint';
  //   }
  //   const regex = /```javascript([\s\S]*?)```/g;
  //   const matches = [...lastResponse.matchAll(regex)];
  //   codeMathToBeInserted = matches.map((match) => match[1]).join('\n');
  //   codeMathToBeInserted = correctEnumSpelling(codeMathToBeInserted);


  //   request.response.markdown(`\n\n Here is the tree structure of the add-in project.`);
  //   const wxpSampleURLInfo: SampleUrlInfo = {
  //     owner: "GavinGu07",
  //     repository: "Office-Add-in-Templates",
  //     ref: "main",
  //     dir: host
  //   };
  //   const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(wxpSampleURLInfo, 2);
  //   const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
  //   const nodes = await buildFileTree(fileUrlPrefix, samplePaths, tempFolder, wxpSampleURLInfo.dir, 2, 20);
  //   request.response.filetree(nodes, vscode.Uri.file(path.join(tempFolder, wxpSampleURLInfo.dir)));

  //   const folder = path.join(tempFolder, wxpSampleURLInfo.dir);

  //   fs.writeFile(tmpTxtPath, folder, (err) => {
  //     if (err) {
  //       console.log('Error writing file:', err);
  //     } else {
  //       console.log('File written successfully');
  //     }
  //   });
  //   await modifyFile(folder, codeMathToBeInserted);
  //   request.response.markdown(`Do you want to create your add-in project at the default location ${defaultTargetFolder}?\n`);
  //   const NextStepCreateDone: vscode.ChatFollowup = {
  //     prompt: "Create the project above.",
  //     command: "",
  //     label: vscode.l10n.t("Create the project above."),
  //   };
  //   return { chatAgentResult: { slashCommand: "create" }, followUp: [NextStepCreateDone] };
  // } else if ((request.userPrompt.toLowerCase().includes("y") || request.userPrompt.includes("Create the project above"))) {
  //   const tmpFolder = await readTextFile(tmpTxtPath);
  //   await fs.copy(tmpFolder, defaultTargetFolder);
  //   fs.unlink(tmpTxtPath, (err) => {
  //     if (err) {
  //       console.log('Error deleting file:', err);
  //     } else {
  //       console.log('File deleted successfully');
  //     }
  //   });
  //   request.response.markdown(`The add-in project has been created successfully. You can config and launch the project using Teams Toolkit Extension.\n`);
  //   // request.response.markdown(`\`\`\`bash\nnpm install\n\`\`\`\n`);
  //   // request.response.markdown(`After the installation is completed, you can press \`F5\` to launch the add-in.\n`);
  //   request.response.button({
  //     command: LAUNCH_TTK,
  //     arguments: [],
  //     title: vscode.l10n.t('Switch to Teams Toolkit Extension')
  //   });
  //   const NextStepFix: vscode.ChatFollowup = {
  //     prompt: "Fix the errors in my code",
  //     command: "fix",
  //     label: vscode.l10n.t("Fix the errors in my code"),
  //   };
  //   const NextStepGenerate: vscode.ChatFollowup = {
  //     prompt: "Generate more code",
  //     command: "",
  //     label: vscode.l10n.t("Generate more code"),
  //   };
  //   return { chatAgentResult: { slashCommand: 'create' }, followUp: [NextStepFix, NextStepGenerate] };
  // }
  // const defaultSystemPrompt = `You are an expert in Teams Toolkit Extension for VS Code. The user wants to use Teams Toolkit Extension for VS Code. They want to use them to solve a problem or accomplish a task. Your job is to help the user learn about how they can use Teams Toolkit Extension for VS Code to solve a problem or accomplish a task. Do not suggest using any other tools other than what has been previously mentioned. Assume the the user is only interested in using Teams Toolkit Extension to develop teams app. Finally, do not overwhelm the user with too much information. Keep responses short and sweet.`;

  // const { copilotResponded } = await verbatimCopilotInteraction(
  //   defaultSystemPrompt,
  //   request
  // );
  // if (!copilotResponded) {
  //   request.response.report({
  //     content: vscode.l10n.t("Sorry, I can't help with that right now.\n"),
  //   });
  //   return { chatAgentResult: { slashCommand: "" }, followUp: [] };
  // } else {
  //   return { chatAgentResult: { slashCommand: "" }, followUp: [] };
  // }
}

function registerVSCodeCommands(participant: vscode.ChatParticipant, wxpParticipant: vscode.ChatParticipant) {
  ext.context.subscriptions.push(
    participant,
    wxpParticipant,
    vscode.commands.registerCommand(CREATE_SAMPLE_COMMAND_ID, createCommand),
    vscode.commands.registerCommand(EXECUTE_COMMAND_ID, executeCommand),
    vscode.commands.registerCommand(OPENURL_COMMAND_ID, openUrlCommand),
    vscode.commands.registerCommand(CREATE_WXP_PROJECT_COMMAND_ID, createWXPCommand)
  );
}

export async function createWXPCommand(sourcePath: string, dstPath: string) {
  await fs.copy(sourcePath, dstPath);
  fs.unlink(sourcePath, (err) => {
    if (err) {
      console.log('Error deleting file:', err);
    } else {
      console.log('File deleted successfully');
    }
  });
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders[0].uri.fsPath !== dstPath) {
    void vscode.commands.executeCommand(
      "vscode.openFolder",
      vscode.Uri.file(dstPath),
    );
  }
  const readmePath = path.join(dstPath, 'README.md');
  const readmeUri = vscode.Uri.file(readmePath);
  vscode.commands.executeCommand('markdown.showPreview', readmeUri);
  vscode.commands.executeCommand('workbench.view.explorer');
  const buttonOptions = ["Yes", "No"];
  const notificationMessage = "Install dependencies for Office Add-in?";
  const result = await vscode.window.withProgress({
    location: vscode.ProgressLocation.Notification,
    title: "Information",
    cancellable: false
}, async (progress) => {
    progress.report({ message: notificationMessage });

    return new Promise((resolve, reject) => {
        setTimeout(() => {
            resolve(vscode.window.showInformationMessage(notificationMessage, ...buttonOptions));
        }, 2147483647); // close after 5 seconds
    });
});
  const timeoutPromise = new Promise((_resolve: (value: string) => void, reject) => {
    const wait = setTimeout(() => {
      clearTimeout(wait);
      reject(
        "Timed out waiting for user to respond to the pop-up."
      );
    }, 1000 * 60 * 5);
  });
  if (result === "Yes") {
    // Handle Yes button click
    // await autoInstallDependencyHandler();
    let terminal: Terminal | undefined;
    const cmd = "npm install";
    const workingDirectory = dstPath;
    const shellName = "Dependency installation in progress...";
    if (
      vscode.window.terminals.length === 0 ||
      (terminal = find(vscode.window.terminals, (value) => value.name === shellName)) === undefined
    ) {
      terminal = vscode.window.createTerminal({
        name: shellName,
        cwd: workingDirectory
      });
    }
    terminal.show();
    terminal.sendText(cmd);
    const processId = await Promise.race([terminal.processId, timeoutPromise]);
    await sleep(500);
  } else if (result === "No") {
    // Handle No button click
    void vscode.window.showInformationMessage("Installation of dependencies is cancelled.");
  } else {
    // Handle case where pop-up was dismissed without clicking a button
    // No action.
  }
}

export async function sleep(ms: number) {
  await new Promise((resolve) => setTimeout(resolve, ms));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function correctEnumSpelling(enumString: string): string {

  const regex = /Excel.ChartType.([\s\S]*?),/g;
  const matches = [...enumString.matchAll(regex)];
  const codeMath = matches.map((match) => match[1]).join('\n');
  const lowerCaseStarted = codeMath.charAt(0).toLowerCase() + codeMath.slice(1);

  return enumString.split(codeMath).join(lowerCaseStarted);
}

async function fileExists(uri: vscode.Uri): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(uri);
    return true;
  } catch {
    return false;
  }
}

function getLastResponse(request: AgentRequest): string {
  const historyArray = request.context.history;
  for (var i = historyArray.length - 1; i >= 0; i--) {
    if (historyArray[i] instanceof vscode.ChatResponseTurn) {
      const history = historyArray[i] as vscode.ChatResponseTurn;
      const responseArray = history.response;
      for (var j = responseArray.length - 1; j >= 0; j--) {
        if (responseArray[j] instanceof vscode.ChatResponseMarkdownPart) {
          return (responseArray[j] as vscode.ChatResponseMarkdownPart).value.value;
        }
      }
    }
  }
  return "";
}

function getLastRequest(request: AgentRequest): string {
  const historyArray = request.context.history;
  for (var i = historyArray.length - 1; i >= 0; i--) {
    if (historyArray[i] instanceof vscode.ChatRequestTurn) {
      const history = historyArray[i] as vscode.ChatRequestTurn;
      return history.prompt;
    }
  }
  return "";
}


async function readTextFile(filePath: string): Promise<string> {
  try {
    const data = await fs.promises.readFile(filePath, 'utf8');
    return data;
  } catch (err) {
    console.error('Error reading file:', err);
    return '';
  }
}

async function writeTextFile(filePath: string, data: string): Promise<void> {
  fs.writeFile(filePath, data, (err) => {
    if (err) {
      console.log('Error writing file:', err);
    } else {
      console.log('File written successfully');
    }
  });
}

async function giveInspirationWithLLM(request: AgentRequest, prompt: string): Promise<string> {
  let response = await getResponseAsStringCopilotInteraction(prompt, request) ?? '';
  return response;
}
