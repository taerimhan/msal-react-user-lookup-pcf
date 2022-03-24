import * as React from "react";
import { createContext, useContext } from "react";
import { IInputs } from "./generated/ManifestTypes";
import { PopupRequest } from "@azure/msal-browser";

export interface IAppContext {
    componentContext: ComponentFramework.Context<IInputs>;
    tokenRequest: PopupRequest
}

export const AppContext = createContext<IAppContext>({} as IAppContext);

export const AppProvider: React.FC<IAppContext> = (props) => {
    const { componentContext, tokenRequest } = props;

    return (
        <AppContext.Provider value={{ componentContext, tokenRequest: tokenRequest }} >
            {props.children}
        </AppContext.Provider>
    );
};

export const useAppContext = () => useContext(AppContext);