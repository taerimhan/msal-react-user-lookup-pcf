import * as React from "react";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useAccount, useMsal } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { ActionButton } from "@fluentui/react/lib/Button";
import { Stack } from "@fluentui/react/lib/Stack";
import { useAppContext } from "../AppContext";

export const Login: React.FC = () => {
    const { instance, inProgress, accounts } = useMsal();
    const account = useAccount(accounts[0] || {});
    const { tokenRequest } = useAppContext();

    return (
        <Stack
            horizontalAlign={"end"}
            horizontal
            tokens={{ childrenGap: 4 }}
            verticalAlign={"center"}>
            <AuthenticatedTemplate>
                <ActionButton
                    iconProps={{ iconName: "Signout" }}
                    onClick={() => instance.logout()}>
                    Sign Out {account?.name ? `(${account.name})` : ''}
                </ActionButton>
            </AuthenticatedTemplate>
            <UnauthenticatedTemplate>
                <ActionButton
                    iconProps={{ iconName: "Signin" }}
                    onClick={() => instance.loginPopup(tokenRequest)}
                    disabled={inProgress === InteractionStatus.Login}>
                    Sign in
                </ActionButton>
            </UnauthenticatedTemplate>
        </Stack>
    );
}