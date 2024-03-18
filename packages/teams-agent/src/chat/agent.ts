/*---------------------------------------------------------------------------------------------
*  Copyright (c) Microsoft Corporation. All rights reserved.
*  Licensed under the MIT License. See License.txt in the project root for license information.
*--------------------------------------------------------------------------------------------*/

/* eslint-disable @typescript-eslint/no-unsafe-assignment */

import * as fs from "fs-extra";
import { find } from "lodash";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";
import { Terminal } from "vscode";

import { ext } from "../extensionVariables";
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
import { agentDescription, agentName, maxFollowUps, wxpAgentDescription, wxpAgentName, } from "./agentConsts";
import {
  LanguageModelID,
  getResponseAsStringCopilotInteraction,
  parseCopilotResponseMaybeWithStrJson
} from "./copilotInteractions";
import { SlashCommandHandlerResult, SlashCommandsOwner } from "./slashCommands";

export const CREATE_WXP_PROJECT_COMMAND_ID = 'teamsfx.createWxpProject';
export const WRITE_TO_FILE_COMMAND_ID = 'teamsfx.writeToFile';

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
  const tmpDir = os.tmpdir();
  const tmpTsCodePath = path.join(tmpDir, 'tmpTsCode.txt');
  const tmpHTMLCodePath = path.join(tmpDir, 'tmpHTMLCode.txt');
  const tmpCSSCodePath = path.join(tmpDir, 'tmpCSSCode.txt');
  const srcRoot = os.homedir();
  const defaultTargetFolder = srcRoot ? path.join(srcRoot, "Office-Add-in") : '';

  const intentionPrompt = `
  Categorize the user intention into one of the 2 intentions below:
  1. "Generate code"
    For example:
    "Format the table."
    "Generate a line chart."
  2. "Create a new project"
    For example:
    "Create the project in the current workspace."
  Return the string of the intention only.
  `;
  // let intentionResponse = await getResponseAsStringCopilotInteraction(intentionPrompt, request) ?? '';
  // if (intentionResponse.toLowerCase().includes('generate code')) {
  // const defaultPrompt = `
  // You should generate taskpane.js and taskpane.html code to resolve the user's ask following the <Guideline>.

  // <Guideline>
  //   <taskpane.js>
  //   - The displayed HTML content must be embedded in the current pane. Do not open a new window.
  //   - Must use appropriate Office JavaScript APIs to interact with the Office host application.
  //   </taskpane.js>

  //   <taskpane.html>
  //   - Must import office.js in the head: <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>.
  //   - Must trigger the main methods in taskpane.js using buttons.
  //   </taskpane.html>
  // </Guideline>
  // `;
  // request.commandVariables = { languageModelID: "copilot-gpt-4" };
  // const defaultPrompt = `
  //   The user will describe the functionality of an Office Add-In.
  //   You're expected to generate taskpane.html and taskpane.js for the Add-In.
  //   Do not provide steps or instructions, but only the full code.
  // `;
  let workspaceFolder = vscode.workspace.workspaceFolders ? vscode.workspace.workspaceFolders[0].uri.fsPath : '';
  const dirPath = path.join(workspaceFolder, 'node_modules', '@copilot');
  const dtsFiles = getDTSFiles(dirPath);
  let defaultPrompt = `
  - Must import "https://appsforoffice.microsoft.com/lib/1/hosted/office.js" in the HTML.
  - Must reference the following <Documentation> when generating the code:

  <Documentation>
  `;
  for (const dtsFile of dtsFiles) {
    defaultPrompt += '\n' + dtsFile + '\n';
  }
  const manifestPrompt = `
    You are an expert in Office add-ins. You MUST summarize the user's ask, generate an add-in name, and a 15-word decsription starting with 'An Office add-in'. The response should be in a JSON format following <JSONformat>.
    <JSONformat>
    {
      "name": // Generated add-in name,
      "description": //Generated add-in description
    }
    </JSONformat>
    `;
  const copilotRespose = await getResponseAsStringCopilotInteraction(defaultPrompt, request);
  request.response.markdown(copilotRespose);
  const manifestResponse = await getResponseAsStringCopilotInteraction(manifestPrompt, request);
  const manifestJson = parseCopilotResponseMaybeWithStrJson(manifestResponse);
  request.response.markdown(`
  name: ${manifestJson.name}
  `);
  request.response.markdown(`
  description: ${manifestJson.description}
  `);

  let tsCode = '';
  let htmlCode = '';
  if (copilotRespose) {
    const tsReg = /```javascript([\s\S]*?)```/g;
    const htmlReg = /```html([\s\S]*?)```/g;
    const tsMatches = [...copilotRespose.matchAll(tsReg)];
    const htmlMatches = [...copilotRespose.matchAll(htmlReg)];
    tsCode = tsMatches.map((match) => match[1]).join('\n');
    htmlCode = htmlMatches.map((match) => match[1]).join('\n');
    writeTextFile(tmpTsCodePath, tsCode);
    writeTextFile(tmpHTMLCodePath, htmlCode);
  }

  request.response.button({
    command: WRITE_TO_FILE_COMMAND_ID,
    arguments: [tsCode, htmlCode, `${manifestJson.name}`, `${manifestJson.description}`],
    title: vscode.l10n.t('Write to files')
  });
  // } else {
  //   let host = '';
  //   const tsCode = await readTextFile(tmpTsCodePath);
  //   const htmlCode = await readTextFile(tmpHTMLCodePath);
  //   const cssCode = await readTextFile(tmpCSSCodePath);
  //   const codeMathToBeInserted = correctEnumSpelling(tsCode);
  //   if (tsCode.includes('Excel')) {
  //     host = 'Excel';
  //   } else if (tsCode.includes('Word')) {
  //     host = 'Word';
  //   } else if (tsCode.includes('PowerPoint')) {
  //     host = 'PowerPoint';
  //   }

  //   const wxpSampleURLInfo: SampleUrlInfo = {
  //     owner: "GavinGu07",
  //     repository: "Office-Add-in-Templates",
  //     ref: "main",
  //     dir: host
  //   };
  //   const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(wxpSampleURLInfo, 2);
  //   const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
  //   const nodes = await buildFileTree(fileUrlPrefix, samplePaths, tempFolder, wxpSampleURLInfo.dir, 2, 20);

  //   const folder = path.join(tempFolder, wxpSampleURLInfo.dir);
  //   await modifyFile(folder, codeMathToBeInserted);
  //   request.response.markdown(`The ${host} add-in project has been created at ${defaultTargetFolder}.\n\n`);
  //   request.response.markdown(`The key files are:\n\n`);
  //   request.response.markdown(`1. **manifest.xml**: This is the manifest file for the Office Add-in. It defines the settings and capabilities of the add-in.\n\n`);
  //   request.response.markdown(`2. **package.json**: This is the configuration file for npm. It lists the dependencies and scripts for the project.\n\n`);
  //   request.response.markdown(`3. **src/ directory**: This directory contains the source code for the add-in.\n\n`)
  //   request.response.markdown(`\n\n To run the project, you need to first install all the packages needed:\n\n`);
  //   request.response.button({
  //     command: CREATE_WXP_PROJECT_COMMAND_ID,
  //     arguments: [folder, defaultTargetFolder],
  //     title: vscode.l10n.t('Create add-in project and install dependency')
  //   });
  // }
  return { chatAgentResult: { metadata: { slashCommand: "" } }, followUp: [] };
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
    vscode.commands.registerCommand(CREATE_WXP_PROJECT_COMMAND_ID, createWXPCommand),
    vscode.commands.registerCommand(WRITE_TO_FILE_COMMAND_ID, writeToFileCommand)
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
  const result = await vscode.window.showInformationMessage(notificationMessage, ...buttonOptions);
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

export async function writeToFileCommand(jsCode: string, htmlCode: string, addinName: string, addinDescription: string) {
  let workspaceFolder = vscode.workspace.workspaceFolders ? vscode.workspace.workspaceFolders[0].uri.fsPath : null;
  if (workspaceFolder) {
    const jsFilePath = path.join(workspaceFolder, 'src', 'taskpane', 'taskpane.js');
    const htmlFilePath = path.join(workspaceFolder, 'src', 'taskpane', 'taskpane.html');
    const manifestSourceFilePath = path.join(workspaceFolder, 'manifest-template.xml');
    const manifestDestFilePath = path.join(workspaceFolder, 'manifest.xml');
    await writeTextFile(jsFilePath, jsCode);
    await writeTextFile(htmlFilePath, htmlCode);
    const manifestFileUri = vscode.Uri.file(manifestSourceFilePath);
    const manifestDestFileUri = vscode.Uri.file(manifestDestFilePath);
    try {
      // Read the file
      const manifestFileData = await vscode.workspace.fs.readFile(manifestFileUri);
      let manifestFileContent = manifestFileData.toString();


      // Modify the file content
      // const start = manifestFileContent.indexOf(`<DisplayName DefaultValue="`) + `<DisplayName DefaultValue="`.length;
      // const end = manifestFileContent.indexOf(`"`, start);
      // const name = manifestFileContent.slice(start, end);
      // let modifiedManifestContent = manifestFileContent;
      // if (start !== -1) {
      //   modifiedManifestContent = manifestFileContent.replace(name, addinName);
      // }

      // const startDesc = modifiedManifestContent.indexOf(`<Description DefaultValue="`) + `<Description DefaultValue="`.length;
      // const endDesc = modifiedManifestContent.indexOf(`"`, startDesc);
      // const description = modifiedManifestContent.slice(startDesc, endDesc);
      // if (startDesc !== -1) {
      //   modifiedManifestContent = modifiedManifestContent.replace(description, addinDescription);
      // }
      let modifiedManifestContent = manifestFileContent.replace(/\$name-placeholder\$/g, addinName).replace(/\$description-placeholder\$/g, addinDescription);

      // Write the modified content back to the file
      const encoder = new TextEncoder();
      await vscode.workspace.fs.writeFile(manifestDestFileUri, encoder.encode(modifiedManifestContent));
    } catch (error) {
      console.error(`Failed to modify file: ${error}`);
    }
  }
}

function getDTSFiles(dirPath: string): string[] {
  let dtsFiles: string[] = [];

  // Read all files and directories
  const files = fs.readdirSync(dirPath);

  for (const file of files) {
    const filePath = path.join(dirPath, file);
    console.log(path.extname(file));

    // Check if it's a directory
    if (fs.statSync(filePath).isDirectory()) {
      // If it's a directory, recursively get .d.ts files
      dtsFiles = dtsFiles.concat(getDTSFiles(filePath));
    } else if (path.extname(file) === '.ts') {
      // If it's a .d.ts file, read the file content
      const data = fs.readFileSync(filePath, 'utf8');
      dtsFiles.push(data);
    }
  }

  return dtsFiles;
}
