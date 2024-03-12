import { AddinSampleNode } from './addInSampleNode';
import { apiSampleData } from './apiSamples'; // Import the content of apiSamples.json
import { SampleData } from './sampleData'; // Import the content of apiSamples.json

export class SampleProvider {
  private rootSample: AddinSampleNode;
  private static instance: SampleProvider;
  private isSampleDataInitialized: boolean = false;

  private constructor() {
    // Private constructor to prevent direct instantiation
    this.rootSample = new AddinSampleNode("root");
  }

  public static getInstance(): SampleProvider {
    if (!SampleProvider.instance) {
      SampleProvider.instance = new SampleProvider();
    }
    return SampleProvider.instance;
  }

  public initSampleData() {
    // Load the sample data from the json file
    apiSampleData.samples.forEach((sample) => {
      let namespace = sample.namespace;
      let name = sample.name;
      let docLink = sample.docLink;
      let code = sample.sample;
      let scenario = sample.scenario;

      this.addSample(namespace.toLowerCase(), name.toLowerCase(), docLink, code, scenario);
    });

    this.isSampleDataInitialized = true;
  }

  addSample(namespace: string, name: string, docLink: string, code: string, scenario: string) {
    // The namespace sometimes presents sevelar levels of nested samples, so we need to split it using "."
    const namespaceArray = namespace.split('.');

    let currentNode: AddinSampleNode = this.rootSample;
    for (let i = 0; i < namespaceArray.length; i++) {
      let nestedNodeName = namespaceArray[i].toLowerCase();
      let nestedNode = currentNode.getNestedSampleNode(nestedNodeName);
      if (!nestedNode) {
        nestedNode = new AddinSampleNode(nestedNodeName);
        currentNode.nestedSampleNode.set(nestedNodeName, nestedNode);
      }
      currentNode = nestedNode;
    }

    currentNode.addSample(name, docLink, code, scenario);
  }

  public getAPISampleCodes(className: string, name: string): Map<string, SampleData> {
    if (!this.isSampleDataInitialized) {
      this.initSampleData();
    }

    // The className sometimes presents sevelar levels of nested samples, so we need to split it using "."
    const pathArray = className.split('.');

    let sampleCandidate: Map<string, SampleData> = new Map();
    let currentNode: AddinSampleNode | undefined = this.rootSample;
    let findNestedNode = true;
    for (let i = 0; i < pathArray.length && findNestedNode; i++) {
      let nestedNodeName = pathArray[i].toLowerCase();
      let nestedNode = currentNode?.getNestedSampleNode(nestedNodeName);
      if (!nestedNode) {
        findNestedNode = false;
        break;
      }
      currentNode = nestedNode;
    }

    if (findNestedNode) {
      let sampleData = currentNode?.getMostRelevantSampleCodes(name);

      if (!!sampleData) {
        sampleCandidate.set(sampleData.name, sampleData);
      }
    }

    return sampleCandidate;
  }
}
