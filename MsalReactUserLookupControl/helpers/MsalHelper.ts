import { AccountInfo, IPublicClientApplication, PopupRequest, InteractionRequiredAuthError } from "@azure/msal-browser";

export const acquireTokenRequest = (instance: IPublicClientApplication, account: AccountInfo | null, tokenRequest: PopupRequest) => {
    if (account) {
        return instance.acquireTokenSilent({
            ...tokenRequest,
            account: account
        }).catch(error => {
            if (error instanceof InteractionRequiredAuthError) {
                return instance.acquireTokenPopup(tokenRequest);
            }
        });
    };
}