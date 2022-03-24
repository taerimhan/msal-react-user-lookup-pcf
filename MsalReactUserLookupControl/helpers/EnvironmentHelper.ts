export interface IEnvironmentVariable {
    type: EnvironmentVariableType;
    value?: any;
}

export type EnvironmentVariableType =
    | "string"
    | "number"
    | "boolean"
    | "json"
    | "unspecified";

export class EnvironmentHelper {
    constructor(private _webAPI: ComponentFramework.WebApi) { }

    async getValue(schemaName: string | null): Promise<any> {
        if (!schemaName) return null;
        const environmentVar = await this.getEnvironmentVariable(schemaName);
        return this.getEnvironmentVariableValue(environmentVar);
    }

    async getEnvironmentVariable(schemaName: string): Promise<IEnvironmentVariable | undefined> {
        const relationshipName = "environmentvariabledefinition_environmentvariablevalue";

        let options = "?";
        options += "$select=schemaname,defaultvalue,type";
        options += `&$filter=statecode eq 0 and schemaname eq '${schemaName}'`;
        options += `&$expand=${relationshipName}($filter=statecode eq 0;$select=value)`;

        const response
            = await this._webAPI.retrieveMultipleRecords("environmentvariabledefinition", options);

        const environmentVarEntity = response.entities.shift();
        if (environmentVarEntity) {
            const environmentVarType
                = this.getEnvironmentVariableType(environmentVarEntity.type);

            const environmentVarValueEntity
                = (<any[]>environmentVarEntity[relationshipName]).shift();

            const environmentVarValue
                = environmentVarValueEntity?.value ?? environmentVarEntity.defaultvalue;

            return {
                type: environmentVarType,
                value: environmentVarValue
            };
        }
    }

    getEnvironmentVariableValue(environmentVar?: IEnvironmentVariable): any {
        if (!environmentVar || !environmentVar.value) {
            return null;
        }

        let value: any = null;
        switch (environmentVar.type) {
            case "string":
                value = this.getStringValue(environmentVar.value);
                break;

            case "boolean":
                value = this.getBooleanValue(environmentVar.value);
                break;

            case "number":
                value = this.getNumberValue(environmentVar.value);
                break;

            case "json":
                value = this.getJsonValue(environmentVar.value);
                break;
        }

        return value;
    }

    private getEnvironmentVariableType(typeValue: number): EnvironmentVariableType {
        let type: EnvironmentVariableType;
        switch (typeValue) {
            case 100000000: type = "string"; break;
            case 100000001: type = "number"; break;
            case 100000002: type = "boolean"; break;
            case 100000003: type = "json"; break;
            default: type = "unspecified"
        }
        return type;
    }

    private getStringValue(rawValue: any): string {
        return rawValue as string;
    }

    private getBooleanValue(rawValue: any): boolean {
        return rawValue == "yes" ? true : false;
    }

    private getNumberValue(rawValue: any): number | null {
        const parsedValue = parseFloat(rawValue);
        if (isNaN(parsedValue)) {
            console.log("Error parsing number value", rawValue);
            return null;
        }
        return parsedValue;
    }

    private getJsonValue(rawValue: any): any {
        try {
            return JSON.parse(rawValue);
        } catch (err) {
            console.log("Error parsing json value", rawValue);
            return null;
        }
    }
}