# RCI Marriage

The RCI Marriage Ministry app for 2027, **Fruit of Our Union**.

- **Demo:** https://astmacary.github.io/rci-marriage/ (sample couple and data, saved on the viewer's device only)
- **Leader view in the demo:** https://astmacary.github.io/rci-marriage/#leaders

It is built as one installable web page (`index.html`) on GitHub Pages. People sign in with Microsoft and the data lives in SharePoint lists. While `CONFIG.clientId` is empty, or the page is opened anywhere other than an address in `CONFIG.appHosts`, it runs in demo mode.

## What's in the app

| Tab | What couples do |
| --- | --- |
| Today | Daily **Reveal Together** prompt, this week's step, what's next |
| Us | **Sunday check-in** (1–5 ratings, two questions, optional mentor heads-up), **Our Orchard**, **Union Goals** |
| Grow | The month's Plant / Water / Grow / Harvest weeks: leader video, devotional, couple challenge |
| Branch | Small group of 4–5 couples, mentor couple, mentor request |
| Gather | Monthly gathering RSVP, Marriage Retreat interest, Harvest Celebration |

The **leader dashboard** shows participation, Branch health, sign-ups, mentor requests and the content schedule. It shows who is taking part, never what couples write.

Content written so far: the May (Patience) screens and the January (The Vine) daily prompts. The other months get added as they are written.

## Live connection

The default GitHub Pages address is the live Microsoft sign-in experience. Use `?demo=1` for the fictional demo. Off-host previews always remain demo-only.

`config.js` contains public identifiers only. It reuses the existing **RCI Productions** single-tenant registration, with the Marriage URL registered as an additional SPA redirect. There is no client secret. The app requests only the existing delegated Microsoft Graph `Sites.ReadWrite.All` and `User.Read` scopes, not the registration's other permissions. SharePoint permissions still determine each signed-in person's access.

Live features in this release:
- Microsoft sign-in and sign-out; tokens are held by MSAL in session storage.
- Published prompts/resources read from Monthly Content.
- Private prompt answers and check-ins saved through Microsoft Graph to their author-only lists.
- Shared notes read and written only on the signed-in couple's permissioned journal folder, with conflict detection.
- Weekly participation by branch and spouse, using only roster and status records. The leader view never queries private writing.
- Union Goals, Our Orchard, encrypted sealed notes and retreat/Harvest RSVPs. See [Pilot setup](PILOT.md) for the private sync job and SharePoint configuration required to activate them.

Live responses are never saved to demo local storage or bundled in the app. A denied or failed Microsoft request stays an error; the app never substitutes sample data or claims an unconfirmed write succeeded.

Automatic sharing of ordinary answers and check-ins and mentor notifications still require separate ministry workflows. Shared journal notes are immediately visible to both spouses; private prompt answers/check-ins remain private until the sharing flows are configured.

### SharePoint schema

Add an optional **AppData** column (Multiple lines of text, **Plain text**, append disabled) to Couples, Monthly Content, Prompt Answers, Sunday Check-ins and Couple Journal. Keep it out of the default list view. `Title` is the record key; `AppData` holds a versioned JSON record. Do not paste real records into this repository.

- Couples: `{ "version": 1, "kind": "couple", "active": true, "memberObjectIds": [], "memberEmails": [], "journalItemId": "<permissioned folder item ID>" }`. Use tenant object IDs when available, especially for guests. Store each couple's real mapping in SharePoint only. Couples get Read access to their own profile item; leaders can read the roster.
- Monthly Content prompt: `{ "version": 1, "kind": "prompt", "date": "2027-01-01", "text": "A ministry-authored prompt" }`.
- Other Monthly Content: `kind` is `video`, `devotional` or `challenge`, with `text` and an optional HTTPS `url`.
- Private answer: `version`, `kind: "prompt-answer"`, `promptId`, `date`, `answer`.
- Check-in: `version`, `kind: "check-in"`, `date`, `ratings` (four values 1–5), `need`, `saw`.
- Couple Journal folder: `{ "version": 1, "entries": [] }`; entries have an ID, date and text. New notes update that folder's fields with `If-Match`, never a new unpermissioned list-root item. The initial shared-notes field supports about 55,000 characters before archiving is needed.

Run `npm ci && npm test` to check the connection's privacy boundaries and failure handling. `npm run vendor` refreshes the pinned Microsoft library after an intentional dependency update. MSAL's license is in `vendor/MSAL-LICENSE.txt`.

### Repository privacy rules

Never commit real couples' identities or answers, keyed Power Automate URLs, or any client secret. Client and tenant IDs are public configuration. Keep workflows and couple provisioning data in Microsoft 365. Do not copy tenant screenshots or exports into the repository.

### 2. SharePoint site: `/sites/RCIMarriage`
A separate site from Productions, so marriage data never sits next to other ministries' data.

| List | Who can read | Notes |
| --- | --- | --- |
| **Couples** | leaders | Couple, spouse 1 and 2 emails, Branch, mentor couple |
| **Monthly Content** | everyone | Prompts by date and type, video links, devotional, challenge. Leaders edit. |
| **Prompt Answers** | the author only | List setting: *Read items that were created by the user*; *Create and edit items that were created by the user* |
| **Couple Journal** | the two spouses only | One folder per couple with its own permissions (about 30 unique scopes in total, not one per answer). Power Automate writes reveals and shared check-ins here. |
| **Sunday Check-ins** | the author only | Same item-level settings as Prompt Answers |
| **Participation** | leaders | Couple, date, kind (answered / checked in / watched). **No content.** The dashboard reads only this. |
| **Mentor Requests** | the author; mentor couple gets a notice | No details are sent, only "please reach out" |
| **RSVPs** | leaders | Gathering, retreat and Harvest sign-ups |

### 3. Power Automate flows (run by a ministry service account)
1. **Reveal Together:** when a Prompt Answer is created, check whether the spouse has answered the same date. If so, copy both answers into the couple's Couple Journal folder and add a Participation row.
2. **Check-in share:** when a check-in is created, copy it into the couple's folder for the spouse and add a Participation row.
3. **Pilot sync:** use the private scheduled job in [Pilot setup](PILOT.md) for encrypted sealed notes, participation, Orchard progress and RSVP projections. Keep sealed rows out of any ordinary Reveal Together flow.
4. **Mentor heads-up:** when a Mentor Request is created, email or Teams-message the mentor couple "<couple> asked you to reach out", with nothing else.
5. **New couple setup:** when a Couples row is added, create the couple's folder, give the two spouses access and send guest invitations.

### Privacy, said plainly
Item-level permissions keep couples' writing away from other couples and from leaders. **Site owners and tenant admins can still technically open any list.** Keep site ownership to one or two named admins plus the flow service account, never give leaders *Design* or *Full Control* on the private lists, and tell couples this in onboarding.

## Updating the app
Edit the relevant app files, bump `APP_VERSION` in `index.html` and `CACHE` in `sw.js` together, then push to `main`. GitHub Pages republishes in about a minute.
