// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Service } from "typedi";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { getLocalizedString } from "../../../common/localizeUtils";
import { CreateApiKeyArgs } from "./interface/createApiKeyArgs";
import { DriverContext } from "../interface/commonArgs";
import { M365TokenProvider, UserError, err, ok } from "@microsoft/teamsfx-api";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { hooks } from "@feathersjs/hooks/lib";
import { InvalidActionInputError } from "../../../error";
import { apiKeyClientSecretReg, logMessageKeys, maxSecretPerApiKey } from "./utilities/constants";
import { OutputEnvironmentVariableUndefinedError } from "../error/outputEnvironmentVariableUndefinedError";
import { AppStudioScopes, GraphScopes } from "../../../common/tools";
import { CreateApiKeyOutputs } from "./interface/createApiKeyOutputs";
import { ApiSecretRegistrationClientSecret } from "./interface/ApiSecretRegistrationClientSecret";
import {
  ApiSecretRegistration,
  ApiSecretRegistrationAppType,
  ApiSecretRegistrationManageableByUserAccessType,
} from "./interface/ApiSecretRegistration";

const actionName = "apiKey/create"; // DO NOT MODIFY the name
const helpLink = "https://aka.ms/teamsfx-actions/apiKey-create";

@Service(actionName) // DO NOT MODIFY the service name
export class CreateApiKeyDriver implements StepDriver {
  description = getLocalizedString(logMessageKeys.description);
  readonly progressTitle = getLocalizedString(logMessageKeys.progessTitle);

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(
    args: CreateApiKeyArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const summaries: string[] = [];
    let outputs: Map<string, string> = new Map<string, string>();

    try {
      context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));
      this.validateArgs(args);

      if (!outputEnvVarNames) {
        throw new OutputEnvironmentVariableUndefinedError(actionName);
      }

      const state = this.loadStateFromEnv(outputEnvVarNames) as CreateApiKeyOutputs;

      const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
        scopes: AppStudioScopes,
      });
      if (appStudioTokenRes.isErr()) {
        throw appStudioTokenRes.error;
      }

      if (state.registrationId) {
        // Registration aleady exists. Will check if registration id exists.
      } else {
        // Registe a new api secret
        const createApiKeyInputs = await this.parseArgs(context.m365TokenProvider, args);

        // TODO: call app studio api
        state.registrationId = "fake-registration-id";

        // TODO: remove test code
        context.logProvider.info("createApiKeyInputs: " + JSON.stringify(createApiKeyInputs));
      }

      outputs = this.mapStateToEnv(state, outputEnvVarNames);
      return {
        result: ok(outputs),
        summaries: summaries,
      };
    } catch (error) {
      // Error handling
      return {
        result: err(error as UserError),
        summaries: summaries,
      };
    }
  }

  private validateArgs(args: CreateApiKeyArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.name !== "string" || !args.name) {
      invalidParameters.push("name");
    }

    if (typeof args.appId !== "string" || !args.appId) {
      invalidParameters.push("appId");
    }

    if (args.clientSecret && !this.validateSecret(args.clientSecret)) {
      invalidParameters.push("apiKey");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters, helpLink);
    }
  }

  // Needs to validate the parameters outside of the function
  private loadStateFromEnv(
    outputEnvVarNames: Map<string, string>
  ): Record<string, string | undefined> {
    const result: Record<string, string | undefined> = {};
    for (const [propertyName, envVarName] of outputEnvVarNames) {
      result[propertyName] = process.env[envVarName];
    }
    return result;
  }

  // Needs to validate the parameters outside of the function
  private mapStateToEnv(
    state: Record<string, string>,
    outputEnvVarNames: Map<string, string>
  ): Map<string, string> {
    const result = new Map<string, string>();
    for (const [outputName, envVarName] of outputEnvVarNames) {
      result.set(envVarName, state[outputName]);
    }
    return result;
  }

  private async parseArgs(
    tokenProvider: M365TokenProvider,
    args: CreateApiKeyArgs
  ): Promise<ApiSecretRegistration> {
    const currentUserRes = await tokenProvider.getJsonObject({ scopes: GraphScopes });
    if (currentUserRes.isErr()) {
      throw currentUserRes.error;
    }
    const currentUser = currentUserRes.value;
    const userId = currentUser["oid"] as string;

    const secrets = this.parseSecret(args.clientSecret!);
    let isPrimary = true;
    const clientSecrets = secrets.map((secret) => {
      const clientSecret: ApiSecretRegistrationClientSecret = {
        value: secret,
        description: args.name,
        priority: isPrimary ? 0 : 1,
        isValueRedacted: true,
      };
      isPrimary = false;
      return clientSecret;
    });

    const apiKey: ApiSecretRegistration = {
      description: args.name,
      targetUrlsShouldStartWith: [],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
      specificAppId: args.appId,
      // targetAudience: "AnyTenant",
      clientSecrets: clientSecrets,
      manageableByUser: [
        {
          userId: userId,
          accessType: ApiSecretRegistrationManageableByUserAccessType.ReadWrite,
        },
      ],
    };

    return apiKey;
  }

  // Allowed inputs: "[secrets1, secrets2]", "secret"
  private parseSecret(apiKeySecret: string): string[] {
    const secrets = apiKeySecret.trim().split(",");
    secrets.map((secret) => secret.trim());
    return secrets;
  }

  private validateSecret(apiKeySecret: string): boolean {
    if (typeof apiKeySecret !== "string") {
      return false;
    }

    const regExp = new RegExp(apiKeyClientSecretReg, "g");
    const regResult = regExp.exec(apiKeySecret);
    if (!regResult) {
      return false;
    }

    const secrets = this.parseSecret(apiKeySecret);
    if (secrets.length > maxSecretPerApiKey) {
      return false;
    }

    return true;
  }
}
