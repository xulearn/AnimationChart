export class Country {
  name: string;
  value: number;
  mapColumn: number;
  increasement: number;
  color: string;

  constructor(name: string, value: number, mapColumn: number, increasement: number, color: string) {
    this.name = name;
    this.value = value;
    this.mapColumn = mapColumn;
    this.increasement = increasement;
    this.color = color;
  }

  setValue(value: number): void {
    this.value = value;
  }

  setIncreasement(increasement: number) {
    this.increasement = increasement;
  }

  setColor(color: string) {
    this.color = color;
  }

  updateIncrease(): void {
    this.value = this.value + this.increasement;
  }
}