import { AgentRequest } from "../../chat/agent";
import { getResponseAsStringCopilotInteraction } from "../../chat/copilotInteractions";
import { SlashCommandHandlerResult } from "../../chat/slashCommands";
import { ISkill } from "./skills/iSkill";
import { SkillsManager } from "./skills/skillsManager";

export class Planner {
  private static instance: Planner;

  private constructor() {
    // Private constructor to prevent direct instantiation
  }

  public static getInstance(): Planner {
    if (!Planner.instance) {
      Planner.instance = new Planner();
    }
    return Planner.instance;
  }

  public async processRequest(request: AgentRequest): Promise<SlashCommandHandlerResult> {
    // request.commandVariables = { languageModelID: "copilot-gpt-3.5-turbo" };
    request.commandVariables = { languageModelID: "copilot-gpt-4" };

    const candidate = await this.getAppropriateSkill(request);

    if (candidate === null) {
      return {
        chatAgentResult: { metadata: { slashCommand: "askOfficeAddin" } },
        followUp: [],
      };
    }

    let result = await candidate["skill"].invoke(candidate["additionalInput"], request);
    result.chatAgentResult.metadata.slashCommand = "askOfficeAddin";
    return result;
  }

  async getAppropriateSkill(request: AgentRequest): Promise<object | null> {
    let defaultSystemPrompt = `
      what is the most accurate description of for the current user's ask? Pick the best one from the task list below, strictly follow this format as your output: [your confidence score as xx%] :: [task description] :: [additional input]. additional input should be empty unless it be explicitly set in the task description. And not need to add any explaination for your answer.
      List of tasks:`;

    SkillsManager.getInstance().getSkillsCapability().forEach((item: object) => {
      defaultSystemPrompt = defaultSystemPrompt.concat(`\r\n* ${item["capability"]}. ${item["promptForAdditionalInput"]}`);
    });

    let startTime = Date.now();
    // Perform the desired operation

    const copilotResponse = await getResponseAsStringCopilotInteraction(
      defaultSystemPrompt,
      request
    );

    let endTime = Date.now();
    let timeDifferenceInSeconds = Math.floor((endTime - startTime) / 1000);
    console.log(`[getAppropriateSkill] Time taken to get response from Copilot: ${timeDifferenceInSeconds} seconds.  The response is: ${copilotResponse}`);

    // Extract the confident score from the response
    const regexForConfidenceScore = /\d+/g;
    const confidenceScoreMatch = copilotResponse.match(regexForConfidenceScore) ?? [0];

    // We are not confident enough to proceed
    if (Number(confidenceScoreMatch[0]) < 80) {
      return null;
    }

    // Extract the capability and AdditionalInput from the response
    const regexForCapability = /(?<=:: ).*/g;
    const capabilityAndAdditionalInputMatch = copilotResponse.match(regexForCapability);

    if (capabilityAndAdditionalInputMatch === null || capabilityAndAdditionalInputMatch.length !== 1) {
      // It should not be null, otherwise it is a bug that we need to fix from the prompt.
      return null;
    }

    const parts = capabilityAndAdditionalInputMatch[0].split("::");
    const capability = parts[0].trim();
    const AdditionalInput = parts.length > 1 ? parts[1].trim() : "";

    const capableSkills: ISkill[] = SkillsManager.getInstance().getCapableSkills(capability, AdditionalInput);

    if (capableSkills.length === 0) {
      return [];
    }

    if (capableSkills.length > 1) {
      // If there are multiple matching skills, something wrong with the design of the skills.
      return [];
    }

    return {
      skill: capableSkills[0],
      userIntention: capability,
      additionalInput: AdditionalInput,
    };
  }
}
