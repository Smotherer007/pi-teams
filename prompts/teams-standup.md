---
description: Draft and post a standup update to a Teams channel
argument-hint: "<Team/Channel>"
---
Help me post a standup update to: $@

1. If no channel was given, ask for one.

2. Read the channel's last few posts with `teams_read_channel` so the update
   fits what the team has been discussing, and so I do not repeat something
   that was already said.

3. Ask me for anything you cannot know: what I did yesterday, what I am doing
   today, and what is blocking me. Do not invent progress, and do not carry
   over items from an old update as if they were new.

4. Draft the update in my voice — short, plain, no headings unless the channel
   uses them. Match the format of the previous standups in that channel.

5. Show me the draft and the exact destination channel. Only after I approve,
   post it with `teams_send_channel_message`.
