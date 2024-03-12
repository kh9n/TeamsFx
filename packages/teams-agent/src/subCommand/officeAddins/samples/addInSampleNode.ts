import { SampleData } from './sampleData'; // Import the content of apiSamples.json

// A sample could contains multiple nested samples and multiple real sample codes for different scenarios
export class AddinSampleNode {
  Name: string;
  // this is for nested samples
  // For example, the Excel is a sample, it Workbooks is it's nested sample
  nestedSampleNode: Map<string, AddinSampleNode>;

  // this is for real sample code
  // The key is the scenario description, the value is the code
  sampleCodes: Map<string, SampleData>;

  constructor(name: string) {
    this.Name = name;
    this.nestedSampleNode = new Map<string, AddinSampleNode>();
    this.sampleCodes = new Map<string, SampleData>();
  }

  public getNestedSampleNode(name: string): AddinSampleNode | undefined {
    return this.nestedSampleNode.get(name);
  }

  public getMostRelevantSampleCodes(name: string): SampleData | undefined {
    return this.sampleCodes.get(name);
  }

  public addSample(name: string, docLink: string, code: string, scenario: string) {
    this.sampleCodes.set(name, { name, docLink, sample: code, scenario });
  }
}
