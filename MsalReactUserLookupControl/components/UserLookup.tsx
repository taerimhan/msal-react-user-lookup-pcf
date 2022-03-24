import * as React from "react";
import { useRef, useState } from "react";
import { AuthenticationResult } from "@azure/msal-browser";
import { useAccount, useMsal } from "@azure/msal-react";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { Persona, PersonaPresence, PersonaSize } from "@fluentui/react/lib/Persona";
import { IStackTokens, Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { ITextFieldStyles, TextField, TextFieldBase } from "@fluentui/react/lib/TextField";
import { useAppContext } from "../AppContext";
import { acquireTokenRequest } from "../helpers/MsalHelper";
import { getUserImage, getUserPresence, getUser } from "../helpers/MsGraphHelper";
import { IUserDetail } from "../interfaces/IUserDetail";

const textFieldStyles: Partial<ITextFieldStyles> = {
    root: {
        minWidth: 100,
        flexGrow: 1
    }
};

const stackTokens: IStackTokens = {
    childrenGap: 16
}

export const UserLookup: React.FC = () => {
    const { instance, accounts } = useMsal();
    const { tokenRequest } = useAppContext();
    const account = useAccount(accounts[0] || {});
    const [userDetail, setUserDetail] = useState<IUserDetail | undefined>();
    const [errorMessage, setErrorMessage] = useState<string>();
    const [processing, setProcessing] = useState<boolean>(false);
    const inputEl = useRef<TextFieldBase>(null);

    const searchUser = () => {
        if (inputEl.current?.value) {
            setProcessing(true);
            resetSearch();

            acquireTokenRequest(instance, account, tokenRequest)
                ?.then((response: AuthenticationResult | undefined) => {
                    if (response) {
                        getUser(response.accessToken, inputEl.current?.value!)
                            .then((user: any) => Promise.all([
                                getUserPresence(response.accessToken, user.id),
                                getUserImage(response.accessToken, user.id)
                            ])
                                .then((results: any[]) => {
                                    const userDetails: IUserDetail = {
                                        displayName: user.displayName,
                                        givenName: user.givenName,
                                        id: user.id,
                                        jobTitle: user.jobTitle,
                                        mail: user.mail,
                                        surname: user.surname
                                    };

                                    if (results && results.length === 2) {
                                        if (results[0]) {
                                            userDetails.activity = results[0]?.activity;
                                            userDetails.availability = results[0]?.availability;
                                        }
                                        if (results[1]) {
                                            userDetails.photo = results[1];
                                        }
                                    }
                                    setUserDetail(userDetails);
                                })
                            )
                            .catch(_ => {
                                setErrorMessage(`User '${inputEl.current?.value!}' could not be retrieved.`);
                                setProcessing(false);
                            })
                            .finally(() => {
                                setProcessing(false);
                            })
                    }
                });
        } else {
            resetSearch();
        }
    }

    const resetSearch = () => {
        setErrorMessage("");
        setUserDetail(undefined);
    }

    const userDetailEl = userDetail
        ? <Persona
            size={PersonaSize.size100}
            text={userDetail.displayName}
            secondaryText={userDetail.jobTitle}
            tertiaryText={userDetail.mail}
            imageUrl={userDetail.photo}
            optionalText={mapPresenceActivity(userDetail.activity)}
            presence={mapPresenceAvailability(userDetail.availability)}
        />
        : processing ? null : <Text>{errorMessage ? errorMessage : "No user to show!"}</Text>

    return (
        <Stack tokens={stackTokens}>
            <Stack horizontal verticalAlign="end" tokens={stackTokens}>
                <TextField
                    componentRef={inputEl}
                    autoComplete="off"
                    label="Lookup User"
                    placeholder="Enter User Principal Name or Object Id"
                    styles={textFieldStyles}
                />
                <PrimaryButton text="Search" onClick={searchUser} allowDisabledFocus />
            </Stack>
            {userDetailEl}
        </Stack>
    );
}

const mapPresenceAvailability = (availability?: string): PersonaPresence => {
    switch (availability) {
        case "Available":
        case "AvailableIdle":
            return PersonaPresence.online;
        case "Away":
        case "BeRightBack":
            return PersonaPresence.away;
        case "Busy":
        case "BusyIdle":
            return PersonaPresence.busy;
        case "DoNotDisturb":
            return PersonaPresence.dnd;
        case "Offline":
            return PersonaPresence.offline;
    }
    return PersonaPresence.none;
}

const mapPresenceActivity = (activity?: string): string => {
    switch (activity) {
        case "BeRightBack": return "Be Right Back";
        case "DoNotDisturb": return "Do Not Disturb";
        case "InACall": return "In A Call";
        case "InAConferenceCall": return "In A Conference Call";
        case "InAMeeting": return "In A Meeting";
        case "OffWork": return "Off Work";
        case "OutOfOffice": return "Out Of Office";
        case "PresenceUnknown": return "Presence Unknown";
        case "UrgentInterruptionsOnly": return "Urgent Interruptions Only";
        default:
            return activity || "";
    }
}
