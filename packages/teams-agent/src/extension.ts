import * as vscode from "vscode";
import { registerChatAgent } from "./chat/agent";
import { registerOfficeAddinChatAgent } from "./chat/officeAddinAgent";
import { ext } from "./extensionVariables";

export function activate(context: vscode.ExtensionContext) {
  ext.context = context;
  registerChatAgent();
  registerOfficeAddinChatAgent();
}

export function deactivate() { }
