export class Conference {
  date: Date;
  name: string;
  email: string;
  title: string;
  description: string;
  target: string;

  toString() {
    return `${this.date}, `
      + `name: ${this.name}, `
      + `email: ${this.email}, `
      + `title: ${this.title}, `
      + `description: ${this.description}, `
      + `target: ${this.target}`
      ;
  }
}