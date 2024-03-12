import { ISkill } from './iSkill'; // Replace this import statement


export class SkillsManager {
  private static instance: SkillsManager;
  skillMap: Map<string, ISkill> = new Map<string, ISkill>();

  private constructor() {
    // Private constructor to prevent direct instantiation
  }

  public static getInstance(): SkillsManager {
    if (!SkillsManager.instance) {
      SkillsManager.instance = new SkillsManager();
    }
    return SkillsManager.instance;
  }

  // Add your class methods and properties here
  public register(skill: ISkill): void {
    const { name } = skill;
    this.skillMap.set(name, skill);
  }

  public getSkillsCapability(): object[] {
    let declaredCapabilities: object[] = [];

    for (const skill of this.skillMap.values()) {
      declaredCapabilities.push({
        capability: skill.capability,
        promptForAdditionalInput: skill.promptForAdditionalInput,
      });
    }

    return declaredCapabilities;
  }

  public getCapableSkills(capability: string, additionalInput: string): ISkill[] {
    const capableSkills: ISkill[] = [];
    this.skillMap.forEach((skill: ISkill) => {
      if (skill.canInvoke(capability, additionalInput)) {
        capableSkills.push(skill);
      }
    });

    return capableSkills;
  }
}
