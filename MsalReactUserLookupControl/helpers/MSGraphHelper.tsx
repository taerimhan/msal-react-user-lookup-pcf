import { ResponseType } from "@microsoft/microsoft-graph-client";
import graph = require("@microsoft/microsoft-graph-client");

export const getUser = async (accessToken: string, upn: string) => {
    const client = await getAuthenticatedClient(accessToken);
    const user = client.api(`/users/${upn}`).get();
    return user;
}

export const getUserPresence = async (accessToken: string, id: string) => {
    const client = await getAuthenticatedClient(accessToken);
    const userPresence = client.api(`/users/${id}/presence`).get();
    return userPresence;
}

export const getUserImage = async (accessToken: string, id: string) => {
    const client = await getAuthenticatedClient(accessToken);
    const response = (await client.api(`/users/${id}/photo/$value`)
        .responseType(ResponseType.RAW)
        .get()) as Response;

    if (response.status === 404 || !response.ok) {
        return null;
    }

    const base64Image = await blobToBase64(await response.blob());
    return base64Image;
}

const getAuthenticatedClient = async (accessToken: string) => {
    return graph.Client.init({
        authProvider: (done: any) => {
            done(null, accessToken);
        }
    });
}

const blobToBase64 = async (blob: Blob): Promise<string> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = reject;
        reader.onload = _ => {
            resolve(reader.result as string);
        };
        reader.readAsDataURL(blob);
    });
}