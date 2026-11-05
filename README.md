# RCI Marriage

The RCI Marriage Ministry app for 2027, **Fruit of Our Union**.

- **Demo:** https://astmacary.github.io/rci-marriage/ (sample couple and data, saved on the viewer's device only)
- **Leader view in the demo:** https://astmacary.github.io/rci-marriage/#leaders

It is built the same way as RCI Productions: one installable web page (`index.html`) on GitHub Pages. People sign in with Microsoft and the data lives in SharePoint lists. While `CONFIG.clientId` is empty, or the page is opened anywhere other than an address in `CONFIG.appHosts`, it runs in demo mode.

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

## Going live on SharePoint (same setup as RCI Productions)

### 1. Microsoft Entra ID app registration
Register a single-page app named **RCI Marriage** in the same tenant as RCI Productions.
- Platform: *Single-page application*. Redirect URI: `https://astmacary.github.io/rci-marriage/` (plus a custom domain later if you add one, e.g. `marriage.therampchurch.com`).
- Delegated permissions: `User.Read`, `User.ReadBasic.All`, `Sites.ReadWrite.All`. Delegated means each person can still only reach what SharePoint lets *them* see.
- Couples outside the church tenant sign in as **guests**, the same way they do for Productions.
- Put the Application (client) ID and Directory (tenant) ID in `CONFIG` in `index.html`.

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
3. **Sealed notes:** the Jan 29 answers are held back and released into the folder on the Harvest Celebration date.
4. **Mentor heads-up:** when a Mentor Request is created, email or Teams-message the mentor couple "<couple> asked you to reach out", with nothing else.
5. **New couple setup:** when a Couples row is added, create the couple's folder, give the two spouses access and send guest invitations.

### Privacy, said plainly
Item-level permissions keep couples' writing away from other couples and from leaders. **Site owners and tenant admins can still technically open any list.** Keep site ownership to one or two named admins plus the flow service account, never give leaders *Design* or *Full Control* on the private lists, and tell couples this in onboarding.

## Updating the app
Edit `index.html`, bump `APP_VERSION` in it and `CACHE` in `sw.js` together, then push to `main`. GitHub Pages republishes in about a minute.
