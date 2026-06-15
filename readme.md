# pst-extractor

[![npm](https://img.shields.io/npm/v/pst-extractor.svg)](https://www.npmjs.com/package/pst-extractor) ![github-issues](https://img.shields.io/github/issues/epfromer/pst-extractor.svg) ![stars](https://img.shields.io/github/stars/epfromer/pst-extractor.svg) ![forks](https://img.shields.io/github/forks/epfromer/pst-extractor.svg) ![](https://david-dm.org/epfromer/pst-extractor/status.svg) ![](https://david-dm.org/epfromer/pst-extractor/dev-status.svg) [![codecov](https://codecov.io/gh/epfromer/pst-extractor/branch/master/graph/badge.svg)](https://codecov.io/gh/epfromer/pst-extractor)

## Project Overview

`pst-extractor` is a TypeScript library for reading Microsoft Outlook and Exchange PST/OST files, plus a local web viewer for reviewing mailbox content in a browser.

The repo has two main parts:

- The root package exports the parsing and mailbox traversal APIs you use from Node.js or other TypeScript projects.
- The `example/` app runs a local PST mail explorer backed by the same library and serves the browser UI used for review, search, tagging, downloads, and audit activity.

The browser app is intentionally local-first:

- Mailboxes live under the project `PST/` folder in a case/search layout such as `PST/Case 1/Search 1/mailbox.pst`.
- Removed mailboxes are moved into `PST/_removed`.
- Auth is local username/password with optional MFA, invite-based onboarding, SMTP sender settings, and an activity log.
- Persistent app data such as users, SMTP settings, review state, and the search index can use MongoDB when configured.
- The audit log stays file-based at `example/logs/activity.log`.
- If the app is exposed behind a reverse proxy, set `PUBLIC_BASE_URL` so invite links point at the public URL.

The library parses PST and OST mailbox data without modifying the source files.

## Features

Extract objects from MS Outlook/Exchange PST files.

This is based off code from https://github.com/rjohnsondev/java-libpst. Thanks to Richard Johnson and Orin Eman.

A spec from Microsoft on the PST file format is at https://msdn.microsoft.com/en-us/library/ff385210(v=office.12).aspx.

Note that this tool does NOT work with corrupt PST files.

## Install

```npm install --save pst-extractor```

or

```yarn add pst-extractor```

## Usage

Start with the example app to browse the mailbox files in the project `PST/` folder. Most of the major objects still have Jest test specs which show how the object attributes can be accessed.

For the current browser experience, build the frontend and then run the example server:

```bash
npm install
npm run build:frontend
cd example
npm start
```

```bash
cd example
npm start
```

or

```bash
cd example
yarn start
```

Open the address printed by the server, sign in, and then pick a case/search scope from the left pane. The viewer reads whichever `.pst` or `.ost` files you place under the project `PST/` folder.

### Web viewer

The `example/` package now starts a local PST mail explorer in your browser.

Put the mailbox files you want to browse in the project `PST/` folder using a case/search structure such as:

```text
PST/
  Case 1/
    Search 1/
      mailbox.pst
```

If you place files directly in `PST/`, they appear under the `PST root` scope as a fallback.

Then run:

```bash
npm install
cd example
npm install
npm start
```

You can set `HOST`, `PORT`, `PUBLIC_BASE_URL`, `MONGODB_URI`, `MONGODB_DB`, `M365_AUTH_BYPASS_IPS`, and `CORS_ALLOWED_ORIGINS` in `example/.env` and the server will load them automatically. A sample file is provided at [example/.env.example](example/.env.example).

The example viewer also has a built-in login screen. By default, sign in with `admin` / `pst-extractor`. You can override that account with `AUTH_USERNAME`, `AUTH_PASSWORD`, and `AUTH_SESSION_TTL_MINUTES` in `example/.env`. The admin user now sends invite links instead of creating passwords directly: open the settings cog, choose **Manage users**, enter a username plus email address, and the server will email a one-time onboarding link. Invitees open that link, set their password, and can optionally enroll MFA with a QR code or manual setup key. After MFA enrollment, recovery codes are shown once and can be downloaded. If the app is behind a reverse proxy, set `PUBLIC_BASE_URL` to the public browser URL so invite links point at the external address instead of the backend host. You can also seed the SMTP sender form with `SMTP_ENABLED`, `SMTP_HOST`, `SMTP_PORT`, `SMTP_SECURE`, `SMTP_USERNAME`, `SMTP_PASSWORD`, `SMTP_FROM_NAME`, `SMTP_FROM_ADDRESS`, and `SMTP_REPLY_TO`; those values are only startup defaults. When `MONGODB_URI` is set, the viewer stores local users and SMTP settings in MongoDB. After signing in as the admin account, open the settings cog in the top bar and choose **Manage users** to invite or delete local viewer accounts, click a user row inside that modal to inspect that user’s activity log, or choose **SMTP settings** to configure and test the sender profile. The settings menu also exposes the global **Activity log** view for recent platform activity. Any signed-in user who has not enabled MFA will see a login-time reminder with **Set up MFA** and **Skip for now**, and can later reopen the same MFA enrollment flow from Settings. User management, SMTP settings, and the activity log are admin-only. Activity log entries are written to `example/logs/activity.log`.

The viewer includes a light/dark mode switcher in the signed-in session bar. The login screen stays light.

Open the address printed by the server, then:

1. Sign in on the auth screen.
2. Pick a case/search scope from the dropdown in the left pane, then pick a `.pst` or `.ost` file from the mailbox list.
3. Browse the folder tree, message list, and reading pane.
4. Search and filter by message metadata, switch between newest-first and folder-order sorting, and page through large folders.
5. Flag or tag mail items from the reading pane for later review.
6. Open a message to inspect headers, HTML or plain-text bodies, attachment metadata, embedded email previews, and transport headers.
7. Download the current message as JSON or EML, or download individual attachments directly from the message detail pane.

Swagger docs are available at `/api/docs`, and the OpenAPI JSON is served at `/api/openapi.json`.

If you set `MONGODB_URI`, review state, local users, and SMTP settings are stored in MongoDB. You can optionally set `MONGODB_DB=pst-extractor` to override the database name. The activity log remains a file-based audit trail in `example/logs/activity.log`.

`M365_AUTH_BYPASS_IPS` is a comma-separated list of IPs that can skip the M365 auth middleware. It defaults to the usual loopback addresses for local testing.

`CORS_ALLOWED_ORIGINS` is a comma-separated allowlist for cross-origin browser access to the `/api` routes. Same-origin requests to the built-in UI do not need to be listed.

Refreshing the tab will reopen the last selected case/search and mailbox from `PST/` while the server process is still running.

A simple script looks like this:

```javascript
import { PSTMessage } from 'pst-extractor';
import { PSTFile } from 'pst-extractor';
import { PSTFolder } from 'pst-extractor';
const resolve = require('path').resolve;

let depth = -1;
let col = 0;

const pstFile = new PSTFile(resolve('/path/to/some/pst.pst'));
console.log(pstFile.getMessageStore().displayName);
processFolder(pstFile.getRootFolder());

/**
 * Walk the folder tree recursively and process emails.
 * @param {PSTFolder} folder
 */
function processFolder(folder: PSTFolder) {
    depth++;

    // the root folder doesn't have a display name
    if (depth > 0) {
        console.log(getDepth(depth) + folder.displayName);
    }

    // go through the folders...
    if (folder.hasSubfolders) {
        let childFolders: PSTFolder[] = folder.getSubFolders();
        for (let childFolder of childFolders) {
            processFolder(childFolder);
        }
    }

    // and now the emails for this folder
    if (folder.contentCount > 0) {
        depth++;
        let email: PSTMessage = folder.getNextChild();
        while (email != null) {
            console.log(getDepth(depth) + 
                'Sender: ' + email.senderName + 
                ', Subject: ' + email.subject);
            email = folder.getNextChild();
        }
        depth--;
    }
    depth--;
}

/**
 * Returns a string with visual indication of depth in tree.
 * @param {number} depth
 * @returns {string}
 */
function getDepth(depth: number): string {
    let sdepth = '';
    if (col > 0) {
        col = 0;
        sdepth += '\n';
    }
    for (let x = 0; x < depth - 1; x++) {
        sdepth += ' | ';
    }
    sdepth += ' |- ';
    return sdepth;
}
```

and will generate the following output:

```text
Personal folders
 |- Top of Personal Folders
 |  |- Deleted Items
 |  |- lokay-m
 |  |  |- MLOKAY (Non-Privileged)
 |  |  |  |- TW-Commercial Group
 |  |  |  |  |- Sender: Lee  Dennis, Subject: New OBA's
 |  |  |  |  |- Sender: Reames Julie, Subject: I/B Link Capacity for November and December 2001
 |  |  |  |  |- Sender: dlsmith@pplweb.com, Subject: West Texas Capacity
 |  |  |  |  |- Sender: Buehler  Craig, Subject: EOL Confirmation -Transwestern Pipeline Company
 |  |  |  |  |- Sender: Buehler  Craig, Subject: EOL Confirmation -Transwestern Pipeline Company
 |  |  |  |  |- Sender: Brostad  Karen, Subject: New PNR points for Transwestern
 |  |  |  |  |- Sender: Frazier  Perry, Subject: RR expansion contracts.
 |  |  |  |  |- Sender: Buehler  Craig, Subject: TW EOLs
 |  |  |  |  |- Sender: Lokay, Subject: Bullets 10/26/01
 |  |  |  |  |- Sender: Frazier  Perry, Subject: Verify RR expansion cr's for ROFR.
 |  |  |  |  |- Sender: Frazier  Perry, Subject: Red Rock Adm. cr # 27698 revisions.
 |  |  |  |  |- Sender: Moore, Subject: TW Weekly Report for October 26, 2001
 |  |  |  |  |- Sender: Cabrera  Reyna, Subject: FW: New PNR points for Transwestern
 |  |  |  |  |- Sender: Lee  Dennis, Subject: TW administrative contract # 27698
 |  |  |  |- Systems
 |  |  |  |  |- Sender: ETS General Announcement/ET&S/Enron@ENRON, Subject: How to prepare your expense report
 |  |  |  |  |- Sender: Carl Carter, Subject: eRequest
 |  |  |  |  |- Sender: Puthigai  Savita, Subject: EnronOnline -Stack Manager Changes
 |  |  |  |  |- Sender: Enron Announcements/Corp/Enron@ENRON, Subject: LIM Software Upgrade Notice
 |  |  |  |  |- Sender: Lee  Dennis, Subject: Sending customer information from ENVISION
 |  |  |  |  |- Sender: Enron Global Technology@ENRON, Subject: Email Retention Policy
 |  |  |  |  |- Sender: Enron Messaging Administration, Subject: Supported Internet Email Addresses
 |  |  |  |- Sent Items
 |  |  |  |  |- Sender: Lokay, Subject: Cirque du Soleil - Dralion
 |  |  |  |  |- Sender: Lokay, Subject: Accepted: Updated: Finalize Transwestern Presentations
 |  |  |  |- Personal
 |  |  |  |  |- Sender: Jim Lokay jimbomania@hotmail.com@ENRON, Subject: Fwd: Enjoy fall in an Alamo midsize car -- just $169 a week!
 |  |  |  |  |- Sender: cbulf, Subject: TRIP INFO
 |  |  |  |  |- Sender: Michelle Lokay, Subject: Check out page 7!
 |  |  |  |  |- Sender: ClickAtHome, Subject: ClickAtHome is Coming Soon!
 |  |  |  |  |- Sender: Lisa Norwood, Subject: contact
 |  |  |  |  |- Sender: Jim Lokay jimboman@bigfoot.com@ENRON, Subject: [Fwd: New email address - again!]
 |  |  |  |  |- Sender: Jim Lokay jimboman@bigfoot.com@ENRON, Subject: file
 |  |  |  |  |- Sender: Enron Announcements/Corp/Enron@ENRON, Subject: An Opportunity to Change Your Electricity Provider
 |  |  |  |  |- Sender: James Lokay jlokay@yahoo.com@ENRON, Subject: Fwd: RE: Party Request-the Galleria - Sat 8/11/2001 4:00 PM
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: links
 |  |  |  |  |- Sender: Clickathome,, Subject: Re: Time Warner Cable question
 |  |  |  |  |- Sender: Ramirez  Pilar, Subject: RE: Good Web
 |  |  |  |  |- Sender: Schoolcraft  Darrell, Subject: Wedding
 |  |  |  |  |- Sender: PostOffice@DILBERT.COM@ENRON, Subject: Dilbert Strip from Jim is waiting for you!
 |  |  |  |  |- Sender: Lokay, Subject: eBay item 576466888 (Ends Apr-10-01 104814 PDT) - !!!!Green Is In!!!!! Erin Bud
 |  |  |  |  |- Sender: Lokay, Subject: eBay item 576853567 (Ends Apr-09-01 134449 PDT) - TY Beanie Buddies ERIN Buddy!
 |  |  |  |  |- Sender: Lokay, Subject: eBay item 577532611 (Ends Apr-09-01 175133 PDT) - VALENTINA BUDDY beanie buddie
 |  |  |  |  |- Sender: Lokay, Subject: eBay item 577242315 (Ends Apr-12-01 175442 PDT) - Ty Beanie Buddy & Baby - Vale
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: scans
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: Re: VALENTINA BUDDY beanie buddies RETIRED  Item #577532611
 |  |  |  |  |- Sender: Lokay, Subject: Job Opportunity?
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: Quickie
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: In case you needed to know...
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: graphic
 |  |  |  |  |- Sender: Hitschel  Bonnie V. BHitschel@tesoropetroleum.com@ENRON, Subject: RE:FIESTA
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: Houston
 |  |  |  |  |- Sender: Fawcett  Jeffery, Subject: California Sing-along
 |  |  |  |  |- Sender: Loe  David DLoe@UtiliCorp.com@ENRON, Subject: FW: ONEOK ski trip
 |  |  |  |  |- Sender: Loe  David DLoe@UtiliCorp.com@ENRON, Subject: FW: ONEOK ski trip
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: unit for sale
 |  |  |  |  |- Sender: Hitschel  Bonnie V. BHitschel@tesoropetroleum.com@ENRON, Subject: FW: Even For You
 |  |  |  |  |- Sender: Jennifer Smith @ENRON, Subject: Learn Technical Analysis, Early Bird Special for Houston Class
 |  |  |  |  |- Sender: Watson  Kimberly, Subject: Lindy's B-day
 |  |  |  |  |- Sender: Ramirez  Pilar, Subject: FW: Breast Cancer
 |  |  |  |  |- Sender: Jim Lokay Lokay@bigfoot.com@ENRON, Subject: unverified news report
 |  |  |  |  |- Sender: Lokay, Subject: Accomplishments for Year 2001, as of 05/10
 |  |  |  |  |- Sender: Lokay, Subject: Rose Rec
 |  |  |  |  |- Sender: Hitschel  Bonnie V. BHitschel@tesoropetroleum.com@ENRON, Subject: Voluntary Spyware
 |  |  |  |  |- Sender: Yee  Danny danny.yee@sap.com@ENRON, Subject: Possible Spreadsheet
 |  |  |  |  |- Sender: eserver@enron.com@ENRON, Subject: <<Concur Expense Document>> - Expense060501
 |  |  |  |  |- Sender: Hitschel  Bonnie V. BHitschel@tesoropetroleum.com@ENRON, Subject: Just for fun
 |  |  |  |  |- Sender: Stevens  Missy, Subject: RE: Enron Efforts to the Flood victims
 |  |  |  |  |- Sender: Stuart  Charla, Subject: RE: OPEN ENROLLMENT FOR ENRON KIDS' CENTER IS UNDERWAY!!!
 |  |  |  |  |- Sender: Goradia, Subject: 
 |  |  |  |  |- Sender: jacammarano@pplweb.com@ENRON, Subject: Invitation to Golf Outing
 |  |  |  |  |- Sender: Horton  Stanley, Subject: GOOD JOB!
 |  |  |  |  |- Sender: tsschuler@pplweb.com@ENRON, Subject: PPL EnergyPlus Golf Outing
 |  |  |  |- Sender: Mailbox - Ftenergy1, Subject: Power Mart '01 - Register online for your FREE exhibition pass!
 |- Search Root
 |- SPAM Search Folder 2
```
## Commonly Used Properties - PSTMessage

Note that this is a subset, and all properties are outlined in the respective object .ts file.

Property | Type | Description | Detailed Doco
--- |:---:|---|---
body | string | Plain text message body. | https://msdn.microsoft.com/en-us/library/office/cc765874.aspx
clientSubmitTime | date | Contains the date and time the message sender submitted a message. | https://technet.microsoft.com/en-us/library/cc839781
displayBCC | string | Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;). | https://msdn.microsoft.com/en-us/library/office/cc815730.aspx
displayCC | string | Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;). | https://msdn.microsoft.com/en-us/library/office/cc765528.aspx
displayTo | string | Contains a list of the display names of the primary (To) message recipients, separated by semicolons (;). | https://msdn.microsoft.com/en-us/library/office/cc839687.aspx
getAttachment | PSTAttachment | Get specific attachment from table using index. | 
hasAttachments | boolean | The message has at least one attachment. | https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
isRead | boolean | The message is marked as having been read. | https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
messageClass | string | Contains a text string that identifies the sender-defined message class, such as IPM.Note. | https://msdn.microsoft.com/en-us/library/office/cc765765.aspx
receivedByName | string | Contains the display name of the messaging user who receives the message. | https://msdn.microsoft.com/en-us/library/office/cc840015.aspx
senderEmailAddress | string | Contains the message sender's e-mail address. | https://msdn.microsoft.com/en-us/library/office/cc839670.aspx
senderName | string | Contains the message sender's display name. | https://msdn.microsoft.com/en-us/library/office/cc815457.aspx
subject | string | Contains the full subject of a message. | https://technet.microsoft.com/en-us/library/cc815720

## Author

Ed Pfromer (epfromer@gmail.com)

## License

MIT © [epfromer](https://github.com/epfromer)
