// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FuncValidation,
  Inputs,
  Platform,
  ProjectSettings,
  Stage,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv from "mocked-env";
import os from "os";
import * as path from "path";
import sinon from "sinon";
import { Container } from "typedi";
import { FeatureFlagName } from "../../src/common/constants";
import { isFeatureFlagEnabled, getRootDirectory } from "../../src/common/tools";
import * as tools from "../../src/common/tools";
import {
  ContextUpgradeError,
  FetchSampleError,
  ProjectFolderExistError,
  ReadFileError,
  TaskNotSupportError,
  WriteFileError,
} from "../../src/core/error";
import { QuestionAppName } from "../../src/core/question";
import {
  getAllSolutionPluginsV2,
  getSolutionPluginByName,
  getSolutionPluginV2ByName,
  SolutionPlugins,
  SolutionPluginsV2,
} from "../../src/core/SolutionPluginContainer";
import { parseTeamsAppTenantId } from "../../src/plugins/solution/fx-solution/v2/utils";
import { randomAppName } from "./utils";
import { executeCommand, tryExecuteCommand } from "../../src/common/cpUtils";
import { TaskDefinition } from "../../src/common/local/taskDefinition";
import { execPowerShell, execShell } from "../../src/common/local/process";
import { isValidProject } from "../../src/common/projectSettingsHelper";
import "../../src/plugins/solution/fx-solution/v2/solution";
describe("Other test case", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });
  it("question: QuestionAppName validation", async () => {
    const inputs: Inputs = { platform: Platform.VSCode };
    let appName = "1234";

    let validRes = await (QuestionAppName.validation as FuncValidation<string>).validFunc(
      appName,
      inputs
    );

    assert.isTrue(
      validRes ===
        "Application name must start with a letter and can only contain letters and digits."
    );

    appName = randomAppName();
    const folder = os.tmpdir();
    sandbox.stub(tools, "getRootDirectory").returns(folder);
    const projectPath = path.resolve(folder, appName);

    sandbox.stub<any, any>(fs, "pathExists").withArgs(projectPath).resolves(true);

    validRes = await (QuestionAppName.validation as FuncValidation<string>).validFunc(
      appName,
      inputs
    );
    assert.isTrue(validRes === `Path exists: ${projectPath}. Select a different application name.`);

    sandbox.restore();
    sandbox.stub<any, any>(fs, "pathExists").withArgs(projectPath).resolves(false);

    validRes = await (QuestionAppName.validation as FuncValidation<string>).validFunc(
      appName,
      inputs
    );
    assert.isTrue(validRes === undefined);
  });

  it("error: ProjectFolderExistError", async () => {
    const error = new ProjectFolderExistError(os.tmpdir());
    assert.isTrue(error.name === "ProjectFolderExistError");
    assert.isTrue(
      error.message === `Path ${os.tmpdir()} already exists. Select a different folder.`
    );
  });

  it("error: WriteFileError", async () => {
    const msg = "file not exist";
    const error = WriteFileError(new Error(msg));
    assert.isTrue(error.name === "WriteFileError");
    assert.isTrue(error.message === msg);
  });

  it("error: ReadFileError", async () => {
    const msg = "file not exist";
    const error = ReadFileError(new Error(msg));
    assert.isTrue(error.name === "ReadFileError");
    assert.isTrue(error.message === msg);
  });

  it("error: TaskNotSupportError", async () => {
    const error = new TaskNotSupportError(Stage.createEnv);
    assert.isTrue(error.name === "TaskNotSupportError");
  });

  it("error: FetchSampleError", async () => {
    const error = new FetchSampleError("hello world app");
    assert.isTrue(error.name === "FetchSampleError");
    assert.isTrue(error.message.includes("hello world app"));
  });

  it("isFeatureFlagEnabled: return true when related environment variable is set to 1 or true", () => {
    const featureFlagName = "FEATURE_FLAG_UNIT_TEST";

    let restore = mockedEnv({
      [featureFlagName]: "1",
    });
    assert.isTrue(isFeatureFlagEnabled(featureFlagName));
    assert.isTrue(isFeatureFlagEnabled(featureFlagName, false)); // default value should be override
    restore();

    restore = mockedEnv({
      [featureFlagName]: "true",
    });
    assert.isTrue(isFeatureFlagEnabled(featureFlagName));
    restore();

    restore = mockedEnv({
      [featureFlagName]: "TruE", // should allow some characters be upper case
    });
    assert.isTrue(isFeatureFlagEnabled(featureFlagName));
    restore();
  });

  it("isFeatureFlagEnabled: return default value when related environment variable is not set", () => {
    const featureFlagName = "FEATURE_FLAG_UNIT_TEST";

    const restore = mockedEnv({
      [featureFlagName]: undefined, // delete it from process.env
    });
    assert.isFalse(isFeatureFlagEnabled(featureFlagName));
    assert.isFalse(isFeatureFlagEnabled(featureFlagName, false));
    assert.isTrue(isFeatureFlagEnabled(featureFlagName, true));
    restore();
  });

  it("isFeatureFlagEnabled: return false when related environment variable is set to non 1 or true value", () => {
    const featureFlagName = "FEATURE_FLAG_UNIT_TEST";

    let restore = mockedEnv({
      [featureFlagName]: "one",
    });
    assert.isFalse(isFeatureFlagEnabled(featureFlagName));
    assert.isFalse(isFeatureFlagEnabled(featureFlagName, true)); // default value should be override
    restore();

    restore = mockedEnv({
      [featureFlagName]: "",
    });
    assert.isFalse(isFeatureFlagEnabled(featureFlagName));
    restore();
  });

  it("SolutionPluginContainer", () => {
    const solutionPluginsV2 = getAllSolutionPluginsV2();
    assert.isTrue(solutionPluginsV2.map((s) => s.name).includes("fx-solution-azure"));
    assert.equal(
      getSolutionPluginV2ByName("fx-solution-azure"),
      Container.get(SolutionPluginsV2.AzureTeamsSolutionV2)
    );
    assert.equal(
      getSolutionPluginByName("fx-solution-azure"),
      Container.get(SolutionPlugins.AzureTeamsSolution)
    );
  });

  it("ContextUpgradeError", async () => {
    const userError = ContextUpgradeError(new Error("11"), true);
    assert.isTrue(userError instanceof UserError);
    const sysError = ContextUpgradeError(new Error("11"), false);
    assert.isTrue(sysError instanceof SystemError);
  });

  it("parseTeamsAppTenantId", async () => {
    const res1 = parseTeamsAppTenantId({ tid: "123" });
    assert.isTrue(res1.isOk());
    const res2 = parseTeamsAppTenantId();
    assert.isTrue(res2.isErr());
    const res3 = parseTeamsAppTenantId({ abd: "123" });
    assert.isTrue(res3.isErr());
  });

  it("getRootDirectory", () => {
    let restore = mockedEnv({
      [FeatureFlagName.rootDirectory]: undefined,
    });

    assert.equal(getRootDirectory(), path.join(os.homedir(), "TeamsApps"));
    restore();

    restore = mockedEnv({
      [FeatureFlagName.rootDirectory]: "",
    });

    assert.equal(getRootDirectory(), path.join(os.homedir(), "TeamsApps"));
    restore();

    restore = mockedEnv({
      [FeatureFlagName.rootDirectory]: os.tmpdir(),
    });

    assert.equal(getRootDirectory(), os.tmpdir());
    restore();

    restore = mockedEnv({
      [FeatureFlagName.rootDirectory]: "${homeDir}/TeamsApps",
    });

    assert.equal(getRootDirectory(), path.join(os.homedir(), "TeamsApps"));
    restore();
  });
  it("executeCommand", async () => {
    {
      try {
        const res = await executeCommand("ls", []);
        assert.isTrue(res !== undefined);
      } catch (e) {}
    }
    {
      try {
        const res = await tryExecuteCommand("ls", []);
        assert.isTrue(res !== undefined);
      } catch (e) {}
    }
    {
      try {
        const res = await execShell("ls");
        assert.isTrue(res !== undefined);
      } catch (e) {}
    }
    {
      try {
        const res = await execPowerShell("ls");
        assert.isTrue(res !== undefined);
      } catch (e) {}
    }
  });
  it("TaskDefinition", async () => {
    const appName = randomAppName();
    const projectPath = path.resolve(os.tmpdir(), appName);
    {
      const res = TaskDefinition.frontendStart(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.backendStart(projectPath, "javascript", "echo", true);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.backendWatch(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.authStart(projectPath, "");
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.botStart(projectPath, "javascript", true);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.ngrokStart(projectPath, true, []);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.frontendInstall(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.backendInstall(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.backendExtensionsInstall(projectPath, "");
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.botInstall(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.spfxInstall(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.gulpCert(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.gulpServe(projectPath);
      assert.isTrue(res !== undefined);
    }
  });
  it("isValidProject: true", async () => {
    const projectSettings: ProjectSettings = {
      appName: "myapp",
      version: "1.0.0",
      projectId: "123",
    };
    sandbox.stub(fs, "readJsonSync").resolves(projectSettings);
    const isValid = isValidProject("aaa");
    assert.isTrue(isValid);
  });
});
