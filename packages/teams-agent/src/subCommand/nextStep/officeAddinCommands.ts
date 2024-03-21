import { AgentRequest } from "../../chat/agent";
import {
  SlashCommand,
  SlashCommandHandlerResult,
} from "../../chat/slashCommands";

const officeAddinNextStepCommandName = "nextstep";

// TODO: Need to implement nextStepHandler
export function getOfficeAddinNextStepCommand(): SlashCommand {
  return [
    officeAddinNextStepCommandName,
    {
      shortDescription: `Use this command to move to the next step anytime.`,
      longDescription: `Type this command without additional descriptions to progress to the next step at any stage of Teams apps development.`,
      intentDescription: "",
      handler: (request: AgentRequest) => officeAddinNextStepHandler(request),
    },
  ];
}

// TODO: Need to implement nextStepHandler
async function officeAddinNextStepHandler(
  request: AgentRequest
): Promise<SlashCommandHandlerResult> {
  return {
    chatAgentResult: { metadata: { slashCommand: "nextstep" } },
    followUp: [],
  };
}
