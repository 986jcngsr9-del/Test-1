---
name: slack-unread-filter
description: Scan recent Slack messages across channels, classify them as relevant or noise, and produce a triage report so you can quickly dismiss what doesn't need your attention. Relevant = direct @mention, team ask, decision you need to weigh in on, or something with meaningful impact on your work. Noise = FYI announcements, social chatter, bot posts, channel traffic with no ask or action required.
---

# Slack Unread Filter Skill

Triage recent Slack messages and surface only what actually needs the user's attention.

## Your Slack Identity

Your user ID is **U08U5ADT50V**. Use this to detect direct mentions (`<@U08U5ADT50V>`) in message text.

## What counts as RELEVANT (keep / needs attention)

- Message contains `<@U08U5ADT50V>` — direct mention
- Message contains a @here or @channel AND the content is an ask, incident, or decision (not just an announcement)
- Message is a direct question or request directed at the team that you'd reasonably be expected to respond to or act on
- Incident, outage, or urgent alert in a channel you own or are responsible for
- Someone asking for review, approval, or input on work you're involved in
- A thread reply to a message you sent

## What counts as NOISE (can safely dismiss)

- Automated bot posts, CI/CD notifications, deployment pings with no failure
- FYI announcements with no ask ("just shipped X", "check out this blog post")
- Social/off-topic messages (kudos, watercooler, random)
- @channel/@here messages that are purely informational with no action required
- Status updates from people you're not working directly with
- Messages in high-volume broadcast channels (e.g. #announcements, #general) with no direct ask

## Steps

1. **Get channels to scan**
   - If the user specified channels: use those
   - Otherwise: ask the user which channels to check, or offer to scan the top channels they're likely active in. Start with a few key ones rather than scanning everything at once.

2. **Read recent messages**
   - Use `slack_read_channel` for each channel
   - Default lookback: messages from the last 8 hours. If user specifies a time window, use that.
   - Use `oldest` parameter with a Unix timestamp: current time minus the lookback period in seconds
   - Fetch up to 50 messages per channel (`limit: 50`)
   - Use `response_format: "concise"` to keep context usage low

3. **Classify each message**
   - Apply the RELEVANT vs NOISE criteria above
   - For borderline cases, lean toward RELEVANT — better to surface something than miss it
   - Note the channel, sender, and a brief reason for classification

4. **Check threads on RELEVANT messages**
   - For relevant messages with thread replies, use `slack_read_thread` to see if the thread resolves the issue or if your input is still needed

5. **Produce the triage report**

   Format the output as:

   ```
   ## Slack Triage — [time window]

   ### Needs Attention (N)
   - **#channel** @sender — [brief summary of what they want/need] — [message link or ts]
   - ...

   ### Safe to Dismiss (N)
   - **#channel** — [one-liner description of noise type, e.g. "CI bot posts (12)", "FYI announcement"]
   - ...

   ### Channels Scanned
   - #channel-name (X messages, Y relevant)
   - ...
   ```

6. **Mark-as-read limitation**
   - The Slack MCP does not expose a mark-as-read API. Inform the user that they'll need to manually mark the noise channels as read in Slack.
   - Optionally suggest: "You can use Slack's 'Mark all as read' (Shift+Esc) after reviewing the Needs Attention items above."

## Usage examples

- `/slack-unread-filter` — scan default channels for last 8 hours
- `/slack-unread-filter last 24 hours` — extend the lookback window
- `/slack-unread-filter #eng-platform #incidents` — scan specific channels
- `/slack-unread-filter` can also be invoked as `/slack-unread-filter` with no args for a quick triage

## Notes

- Prioritize speed: scan in parallel where possible (multiple `slack_read_channel` calls at once)
- If a channel has 0 messages in the window, skip it silently
- If the user wants to customize what counts as relevant (e.g. "also flag anything mentioning Project X"), capture that preference and apply it for the session
