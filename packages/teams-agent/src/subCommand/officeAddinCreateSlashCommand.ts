import * as childProcess from "child_process";
import * as fs from "fs-extra";
import * as path from "path";
import * as tmp from "tmp";
import { promisify } from "util";
import * as vscode from "vscode";
import { AgentRequest } from "../chat/agent";
import { SlashCommand, SlashCommandHandlerResult } from "../chat/slashCommands";
import { ProjectMetadata } from "../projectMatch";
import { SampleUrlInfo } from "../sample";
import { buildFileTree, getSampleFileInfo } from "../util";
import { getSampleDownloadUrlInfo } from "./createSlashCommand";

const createOfficeAddinCommandName = "create";
export const CREATE_OFFICEADDIN_SAMPLE_COMMAND_ID = 'teamsAgent.createOfficeAddinSample';

export function getOfficeAddinCreateCommand(): SlashCommand {
  return [createOfficeAddinCommandName,
    {
      shortDescription: `Describe the app you want to build for Microsoft Teams`,
      longDescription: `Describe the app you want to build for Microsoft Teams`,
      intentDescription: '',
      handler: (request: AgentRequest) => officeAddinCreateHandler(request)
    }];
}

// TODO: Implement the create handler function
async function officeAddinCreateHandler(request: AgentRequest): Promise<SlashCommandHandlerResult> {
  request.response.markdown("This is the createOfficeAddinCommandName command");
  return {
    chatAgentResult: { metadata: { slashCommand: "" } },
    followUp: [],
  };
}

async function showOfficeAddinFileTree(projectMetadata: ProjectMetadata | undefined, request: AgentRequest, isCustomFunction: boolean, host: string): Promise<string> {
  request.response.markdown(vscode.l10n.t('\nHere is the files of the sample project.'));
  let downloadUrlInfo: SampleUrlInfo;
  if (projectMetadata) {
    downloadUrlInfo = await getSampleDownloadUrlInfo(projectMetadata.id);
  } else {
    downloadUrlInfo = {
      owner: "OfficeDev",
      repository: isCustomFunction ? "Excel-Custom-Functions" : "Office-Addin-TaskPane-JS",
      ref: "master",
      dir: "",
    };
  }

  const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(downloadUrlInfo, 2);
  const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
  const nodes = await buildFileTree(fileUrlPrefix, samplePaths, tempFolder, downloadUrlInfo.dir, 2, 20);
  const workingDir = process.cwd();
  if (!projectMetadata) {
    try {
      process.chdir(tempFolder);
      await childProcessExec(
        `npm run convert-to-single-host --if-present -- ${host.toLowerCase()}`
      );
      process.chdir(workingDir);
    } catch (error) {
      process.chdir(workingDir);
      console.error('Error downloading files:', error);
      vscode.window.showErrorMessage('Project cannot be downloaded.');
    }
  }
  request.response.filetree(nodes, vscode.Uri.file(path.join(tempFolder, downloadUrlInfo.dir)));
  return path.join(tempFolder, downloadUrlInfo.dir);
}

export async function createOfficeAddinCommand(folder: string) {
  let dstPath = "";
  let folderChoice: string | undefined = undefined;
  if (vscode.workspace.workspaceFolders !== undefined && vscode.workspace.workspaceFolders.length > 0) {
    folderChoice = await vscode.window.showQuickPick(["Current workspace", "Browse..."]);
    if (!folderChoice) {
      return;
    }
    if (folderChoice === "Current workspace") {
      dstPath = vscode.workspace.workspaceFolders[0].uri.fsPath;
    }
  }
  if (dstPath === "") {
    const customFolder = await vscode.window.showOpenDialog({
      title: "Choose where to save your project",
      openLabel: "Select Folder",
      canSelectFiles: false,
      canSelectFolders: true,
      canSelectMany: false,
    });
    if (!customFolder) {
      return;
    }
    dstPath = customFolder[0].fsPath;
  }
  try {
    await fs.copy(folder, dstPath);
    if (folderChoice !== "Current workspace") {
      void vscode.commands.executeCommand(
        "vscode.openFolder",
        vscode.Uri.file(dstPath),
      );
    } else {
      vscode.window.showInformationMessage('Project is created in current workspace.');
      vscode.commands.executeCommand('workbench.view.extension.teamsfx');
    }
  }
  catch (error) {
    console.error('Error creating files:', error);
    vscode.window.showErrorMessage('Project cannot be created.');
  }
}

async function childProcessExec(cmdLine: string): Promise<{ stdout: string; stderr: string }> {
  return promisify(childProcess.exec)(cmdLine);
}
