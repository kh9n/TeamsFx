import * as vscode from 'vscode'; // Add this import statement
import { AgentRequest, ITeamsChatAgentResult } from "../../../chat/agent";
import { SlashCommandHandlerResult } from '../../../chat/slashCommands';

export interface ISkill {
  name: string;
  capability: string;
  promptForAdditionalInput: string;
  canInvoke: (capability: string, additionalInput: string) => boolean;
  invoke: (additionalInput: string, request: AgentRequest) => Promise<SlashCommandHandlerResult>;
  followUp?: (
    result: ITeamsChatAgentResult,
    context: vscode.ChatContext,
    token: vscode.CancellationToken
  ) => vscode.ChatFollowup[] | undefined;
}
