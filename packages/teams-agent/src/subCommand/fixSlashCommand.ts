import * as vscode from 'vscode';
import { type AgentRequest } from "../chat/agent";
import { getResponseAsStringCopilotInteraction } from "../chat/copilotInteractions";
import { type SlashCommand, type SlashCommandHandlerResult } from "../chat/slashCommands";

export const fixCommandName = "fix";

export function getFixCommand(): SlashCommand {
  return [fixCommandName,
    {
      shortDescription: `Describe what kind of app you want to create in Teams`,
      longDescription: `Describe what kind of app you want to create in Teams`,
      intentDescription: '',
      handler: (request: AgentRequest) => fixHandler(request)
    }];
}

async function fixHandler(
  request: AgentRequest
): Promise<SlashCommandHandlerResult> {
  const activeTextEditor = vscode.window.activeTextEditor;
  if (activeTextEditor) {
    let uri = activeTextEditor.document.uri;
    let diagnostics = vscode.languages.getDiagnostics(uri);
    let errorDiagnostics = diagnostics.filter(diagnostic => diagnostic.severity === vscode.DiagnosticSeverity.Error);
    // errorDiagnostics.forEach(diagnostic => {
    //   let errorCode = activeTextEditor.document.lineAt(diagnostic.range.start.line).text;
    //   request.response.markdown(`
    //   - ${diagnostic.message}
    //   - ${errorCode}`);
    // });
    let diagnostic = errorDiagnostics[0];
    if (diagnostic) {
      let errorMessage = diagnostic.message;
      let errorCode = activeTextEditor.document.lineAt(diagnostic.range.start.line).text;
      await fixErrorCode(errorMessage, errorCode, request);
    }
    else {
      request.response.markdown(`Your code seems all good.`);
    }
  }

  return { chatAgentResult: { metadata: { slashCommand: '' } }, followUp: [] };
}

export async function fixErrorCode(errorMessage: string, errorCode: string, request: AgentRequest) {
  const fixErrorCodePrompt = `
  # Role
  I want you act as an expert in Office JavaScript add-in development area.You are also an advisor for Office add-in developers.

  # Instructions
  - There is an error message: ${errorMessage} in the code ${errorCode}. Please give out the right code to fix the error.
  ` ;

  const fixResponse = await getResponseAsStringCopilotInteraction(fixErrorCodePrompt, request);
  if (fixResponse) {
    request.response.markdown(fixResponse);
  }
}
