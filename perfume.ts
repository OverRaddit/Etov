export class Perfume {
    code: string;
    brandName: string;
    name: string;
    keywords: string[];
    accords: string[];

    constructor(code: string, brandName: string, name: string, keywords: string[]) {
        this.code = code;
        this.brandName = brandName;
        this.name = name;
        this.keywords = keywords;
        this.accords = [];
    }
}
