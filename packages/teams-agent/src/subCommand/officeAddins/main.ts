import { CodeGenerator } from './skills/codeGenerator'; // Add the missing import statement
import { SkillsManager } from './skills/skillsManager'; // Add the missing import statement

export function askOfficeAddinInitialize() {
  const skillsManager = SkillsManager.getInstance();
  skillsManager.register(new CodeGenerator());
}
