/// <reference types="@workadventure/iframe-api-typings" />

import axios, { AxiosInstance } from "axios";
import { z } from "zod";

const SSOToken = z.object({
    oid: z.string(),
});

type SSOToken = z.infer<typeof SSOToken>;

const TeamsAvailability = z.enum([
    "Available",
    "AvailableIdle",
    "Away",
    "BeRightBack",
    "Busy",
    "BusyIdle",
    "DoNotDisturb",
    "Offline",
    "PresenceUnknown"
]);

type TeamsAvailability = z.infer<typeof TeamsAvailability>;

const TeamsActivity = z.enum([
    "Available",
    "Away",
    "BeRightBack",
    "Busy",
    "DoNotDisturb",
    "InACall",
    "InAConferenceCall",
    "Inactive",
    "InAMeeting",
    "Offline",
    "OffWork",
    "OutOfOffice",
    "PresenceUnknown",
    "Presenting",
    "UrgentInterruptionsOnly"
]);

type TeamsActivity = z.infer<typeof TeamsActivity>;

console.log('Script started successfully');

let teamsStatusCheckInterval: NodeJS.Timer|undefined = undefined;

WA.onInit().then(() => {
    const metadataIsOk = z.object({
        player: z.object({
            accessTokens: z.array(z.object({
                token: z.string()
            }))
        })
    }).safeParse(WA.metadata);

    console.log("metadata", WA.metadata);

    if (metadataIsOk.success && metadataIsOk.data.player.accessTokens.length > 0) {
        const userToken = metadataIsOk.data.player.accessTokens[0].token;
        let tokenData: SSOToken;

        try {
            tokenData = parseJwt(userToken);
        } catch (err) {
            throw new Error("Your token is not valid");
        }

        const msClient = axios.create({
            baseURL: 'https://graph.microsoft.com/v1.0',
            headers: {
                "Authorization": `Bearer ${userToken}`,
                "Content-Type": "application/json"
            }
        });

        startTeamsStatusCheckInterval(msClient, tokenData);

        WA.player.proximityMeeting.onJoin().subscribe(() => {
            clearInterval(teamsStatusCheckInterval);
            setTeamsStatus(msClient, tokenData, "Busy", "Busy");
        });

        WA.player.proximityMeeting.onLeave().subscribe(() => {
            setTeamsStatus(msClient, tokenData, "Available", "Available");
            startTeamsStatusCheckInterval(msClient, tokenData);
        });
    } else {
        console.info("Your are not connected with the KPMG SSO");
    }
}).catch(e => console.error(e));

/**
 * Parse the JWT token
 * @param token
 * @returns
 */
function parseJwt(token: string): SSOToken {
    var base64Url = token.split('.')[1];
    var base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    var jsonPayload = decodeURIComponent(window.atob(base64).split('').map(function(c) {
        return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
    }).join(''));

    return SSOToken.parse(JSON.parse(jsonPayload));
}

function startTeamsStatusCheckInterval(client: AxiosInstance, token: SSOToken) {
    teamsStatusCheckInterval = setInterval(() => {
        checkTeamsStatus(client, token);
    }, 5000);
}

/**
 * Check the Teams status of the user
 * @param client
 * @param token
 */
function checkTeamsStatus(client: AxiosInstance, token: SSOToken) {
    client.get('/me/presence').then(response => {
        const userPresence = response.data;
        const isPresenceData = z.object({
            availability: TeamsAvailability,
            activity: TeamsActivity,
        }).safeParse(userPresence);

        if (!isPresenceData.success) {
            throw new Error("Your presence status cannot be handled by this script");
        }

        const notAvailable: TeamsAvailability[] = ["BeRightBack", "Busy", "DoNotDisturb"];

        if (notAvailable.includes(isPresenceData.data.availability)) {
            WA.controls.disablePlayerProximityMeeting();
            return;
        }

        const available = ["Available", "Away", "Offline"];

        if (available.includes(isPresenceData.data.availability)) {
            WA.controls.restorePlayerProximityMeeting();

            if (isPresenceData.data.availability === "Offline") {
                setTeamsStatus(client, token, "Available", "Available", "PT2M");
            }
        }
    }).catch(e => console.error("Error while getting Teams status", e));
}

/**
 * Set the Teams status of the user
 * @param client
 * @param token
 * @param availability
 * @param activity
 * @param duration
 */
function setTeamsStatus(
        client: AxiosInstance,
        token: SSOToken,
        availability: Exclude<TeamsAvailability, "AvailableIdle" | "BusyIdle" | "PresenceUnknown">,
        activity: Exclude<TeamsActivity,
            "InACall" |
            "InAConferenceCall" |
            "Inactive" |
            "InAMeeting" |
            "OffWork" |
            "OutOfOffice" |
            "Presenting" |
            "UrgentInterruptionsOnly"
        >,
        duration?: string
) {
    client.post(`/users/${token.oid}/presence/setUserPreferredPresence`, {
        availability: availability,
        activity: activity,
        expirationDuration: duration,
    }).then(() => {
        console.info(`Your presence status has been set to ${availability} - ${activity}`);
    }).catch(e => console.error(e));
}

export {}
