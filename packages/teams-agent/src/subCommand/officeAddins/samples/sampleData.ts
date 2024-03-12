export class SampleData {
  docLink: string;
  sample: string;
  scenario: string;
  name: string;

  constructor(name: string, docLink: string, sample: string, scenario: string) {
    this.docLink = docLink;
    this.sample = sample;
    this.scenario = scenario;
    this.name = name;
  }
}
