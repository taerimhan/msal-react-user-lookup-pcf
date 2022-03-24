import React = require("react");
import ReactDOM = require("react-dom");
import { Configuration, PopupRequest, PublicClientApplication } from "@azure/msal-browser";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { App, IAppProps } from "./App";
import { EnvironmentHelper } from "./helpers/EnvironmentHelper";
import { IConfig } from "./interfaces/IConfig";
const AsyncLock = require('async-lock');

export class MSGraphUserLookupControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _container: HTMLDivElement;
	private _msalInstance: PublicClientApplication;
	private _tokenRequest: PopupRequest;
	private _environmentHelper: EnvironmentHelper;
	private _props: IAppProps;
	private _lock: any = new AsyncLock();

	constructor() { }

	public init(context: ComponentFramework.Context<IInputs>,
		notifyOutputChanged: () => void,
		state: ComponentFramework.Dictionary,
		container: HTMLDivElement) {
		this._container = container;
		this._environmentHelper = new EnvironmentHelper(context.webAPI);
	}

	public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {
		this._lock.acquire("init", async () => {
			if (this._msalInstance == null ||
				context.updatedProperties.includes("env_msalConfig")) {
				this._msalInstance =
					await this.getMsalConfig(context.parameters.env_msalConfig.raw!);
			}

			if (this._tokenRequest == null ||
				context.updatedProperties.includes("env_scopes")) {
				this._tokenRequest = {
					scopes: await this.getScopes(context.parameters.env_scopes.raw || "")
				};
			}
		}).then(() => {
			if (this._msalInstance) {
				this._props = {
					componentContext: context,
					msalInstance: this._msalInstance,
					tokenRequest: this._tokenRequest
				};
				ReactDOM.render(
					React.createElement(App, this._props),
					this._container
				);
			}
		});
	}

	public getOutputs(): IOutputs {
		return {};
	}

	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this._container);
	}

	private getMsalConfig = async (envVarName: string) => {
		const config: IConfig = await this._environmentHelper.getValue(envVarName);
		return new PublicClientApplication({
			auth: {
				clientId: config.clientId,
				authority: config.authority || "https://login.microsoftonline.com/common",
				redirectUri: config.redirectUri || window.location.href,
				postLogoutRedirectUri: config.postLogoutRedirectUri || window.location.href
			}
		} as Configuration);
	}

	private getScopes = async (envVarName: string) => {
		let scopes: string[] = [];
		if (envVarName) {
			scopes = (<string>(await this._environmentHelper.getValue(envVarName)))
				.split(",")
				.map(s => s.trim());
		}
		return scopes;
	}
}