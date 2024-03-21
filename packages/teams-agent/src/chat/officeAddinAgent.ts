import * as vscode from "vscode";
import { ext } from "../extensionVariables";
import {
  helpCommandName
} from "../subCommand/helpSlashCommand";
import { DefaultNextStep } from "../subCommand/nextStep/command";
import { getOfficeAddinNextStepCommand } from "../subCommand/nextStep/officeAddinCommands";
import { CREATE_OFFICEADDIN_SAMPLE_COMMAND_ID, createOfficeAddinCommand, getOfficeAddinCreateCommand } from "../subCommand/officeAddinCreateSlashCommand";
import { AgentRequest, ITeamsChatAgentResult } from "./agent";
import { officeAddinAgentDescription, officeAddinAgentName } from "./agentConsts";
import { SlashCommandHandlerResult, SlashCommandsOwner } from "./slashCommands";

/**
 * Owns slash commands for OfficeAddin agent.
 */
const officeAddinAgentSlashCommandsOwner = new SlashCommandsOwner(
  {
    noInput: helpCommandName,
    default: officeAddinDefaultHandler
  },
  { disableIntentDetection: true }
);
officeAddinAgentSlashCommandsOwner.addInvokeableSlashCommands(
  new Map([
    getOfficeAddinCreateCommand(),
    getOfficeAddinNextStepCommand()
  ])
);

export function registerOfficeAddinChatAgent() {
  try {
    const participant = vscode.chat.createChatParticipant(officeAddinAgentName, officeAddinHandler);
    participant.description = officeAddinAgentDescription;
    participant.iconPath = vscode.Uri.joinPath(
      ext.context.extensionUri,
      "resources",
      "teams.png"
    );
    participant.followupProvider = { provideFollowups: followUpProvider };
    registerVSCodeCommands(participant);
  } catch (e) {
    console.log(e);
  }
}

// TODO: Implement the handler function
async function officeAddinHandler(
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

  const handlers = [officeAddinAgentSlashCommandsOwner];
  for (const handler of handlers) {
    handleResult = await handler.handleRequestOrPrompt(agentRequest);
    if (handleResult !== undefined) {
      break;
    }
  }

  if (handleResult !== undefined) {
    handleResult.followUp = handleResult.followUp?.slice(0, 3);
    return handleResult.chatAgentResult;
  } else {
    return undefined;
  }
}

// TODO: Implement the followUpProvider function
function followUpProvider(
  result: ITeamsChatAgentResult,
  context: vscode.ChatContext,
  token: vscode.CancellationToken
): vscode.ProviderResult<vscode.ChatFollowup[]> {
  const providers = [officeAddinAgentSlashCommandsOwner];

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

// TODO: Implement default handler
async function officeAddinDefaultHandler(
  request: AgentRequest
): Promise<SlashCommandHandlerResult> {
  return {
    chatAgentResult: { metadata: { slashCommand: "" } },
    followUp: [],
  };
}

function registerVSCodeCommands(participant: vscode.ChatParticipant) {
  ext.context.subscriptions.push(
    participant,
    vscode.commands.registerCommand(CREATE_OFFICEADDIN_SAMPLE_COMMAND_ID, createOfficeAddinCommand)
  );
}
