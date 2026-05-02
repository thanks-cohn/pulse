# Pulse

**A tiny command-line life logger**

---

## What is Pulse?

Pulse is a simple tool that lets you log your life in real time using your terminal.

No dashboards.
No friction.
No “open an app.”

You just type what’s happening:

```bash
pulse lunch
pulse wake
pulse pain "left knee hurts"
pulse mood "been sad lately"
```

And Pulse records it as structured JSON.

---

## Philosophy

Pulse is not a productivity app.

It does not try to fix you.
It does not tell you what to do.

It just answers:

```text
What actually happened today?
```

You run it.
You live your life.
Then you look back and enjoy. 

---

------------------------------------------------------------
SETUP + BUILD (LINUX)
------------------------------------------------------------

1) Install dependencies

   Arch / Garuda:
       sudo pacman -S git dotnet-sdk

   Ubuntu / Debian:
       sudo apt update
       sudo apt install git dotnet-sdk-8.0

------------------------------------------------------------
2) Get the project

       git clone https://github.com/thanks-cohn/pulse.git
       cd pulse

------------------------------------------------------------
3) Build

       dotnet build

------------------------------------------------------------
4) Run (quick test)

       dotnet run

------------------------------------------------------------
5) (Optional) Install globally

   Create a standalone binary:

       dotnet publish -c Release -r linux-x64 \
         --self-contained true /p:PublishSingleFile=true

   Then install it:

       sudo cp bin/Release/*/linux-x64/publish/Pulse /usr/local/bin/pulse
       sudo chmod +x /usr/local/bin/pulse

   Now you can run it anywhere:

       pulse


------------------------------------------------------------
WINDOWS — QUICK INSTALL
------------------------------------------------------------

1) Download

Grab the file:

    pulse.exe

(From the Releases page)

------------------------------------------------------------
2) Put it somewhere simple

Create a tools folder:

    C:\tools

Move the file there:

    C:\tools\pulse.exe

------------------------------------------------------------
3) Add it to PATH (one-time setup)

Open Start Menu → search:

    Environment Variables

Click:

    Edit the system environment variables
    → Environment Variables

Under "User variables":

    Path → Edit → New

Add:

    C:\tools

Click OK on everything.

------------------------------------------------------------
4) Restart terminal

Close PowerShell or CMD.
Open it again.

------------------------------------------------------------
5) Test it

    pulse

If you see Pulse respond, you're good.

------------------------------------------------------------
EXAMPLES
------------------------------------------------------------

    pulse mood "good day" tag:softwin
    pulse pain "neck feels better" tag:body
    pulse important "installed Pulse on Windows" tag:tinyvictory
    pulse year 2026

------------------------------------------------------------
NOTES
------------------------------------------------------------

• No installer needed
• No extra dependencies
• Just one file

Pulse is now part of your system.

----------------------------------------------------------
## Core Idea

Every command follows this pattern:

```bash
pulse <event> [time] ["note"]
```

Examples:

```bash
pulse wake
pulse wake "dream journal: ocean, black dog"
pulse wake 9:30am "dream journal: old house"
pulse lunch
pulse lunch "tacos and coffee"
pulse pain "left knee hurts going upstairs"
pulse mood "sad but steady"
```

---

## Events

### Toggle events (start / end automatically)

These commands toggle between start and end:

```bash
pulse meal
pulse breakfast
pulse lunch
pulse dinner
pulse snack
pulse break
pulse walk
pulse work
```

Example:

```bash
pulse lunch "tacos and coffee"
```

→ starts lunch

```bash
pulse lunch
```

→ ends lunch

---

### Point events (single log)

These just log once:

```bash
pulse wake
pulse pain
pulse mood
pulse energy
pulse note
pulse dream
```

Example:

```bash
pulse pain "left knee hurts"
pulse mood "low but functional"
pulse dream "back in school again"
```

---

## Sleep

Sleep is special.

### End your day:

```bash
pulse sleep
```

This logs a sleep attempt.

You can configure it to shut down your machine:

```bash
pulse config sleep-action shutdown
```

Then:

```bash
pulse sleep
```

→ logs event
→ flushes to disk
→ powers off machine

---

### Log actual sleep time:

```bash
pulse sleep 2:40am "couldn't sleep because my mind kept running"
```

This records when you actually fell asleep.

---

## Reports

### Day view

```bash
pulse day
```

Shows everything that happened today.

---

```bash 
pulse day <date>
```

You can look at previous whole days!
(Just use the standard date format YYYY-MM-DD)
Example

```bash 
pulse day 2026-05-02
```


```bash
pulse yeterday
```

You can review recent notes or day events!!






### Recent events

```bash
pulse log
```

---

### Last event (raw JSON)

```bash
pulse last
```

---

### File locations

```bash
pulse where
```

Outputs:

```text
config: ~/.config/pulse/config.json
events: ~/.local/share/pulse/events.jsonl
state:  ~/.local/share/pulse/state.json
```

---

## Editing

Edit last event:

```bash
pulse edit last "better note"
```

Append to last event:

```bash
pulse append last "extra detail"
```

---

## Installation

### Build

```bash
dotnet publish -c Release -r linux-x64 --self-contained true /p:PublishSingleFile=true
```

---

### Install

```bash
mkdir -p ~/.local/bin
cp bin/Release/net*/linux-x64/publish/Pulse ~/.local/bin/pulse
chmod +x ~/.local/bin/pulse
```

---

## Data Format

Pulse stores everything as JSON lines:

```text
~/.local/share/pulse/events.jsonl
```

Each line is a valid JSON object:

```json
{"id":"abc123","ts":"2026-04-30T12:44:00-05:00","type":"lunch_start","category":"meal","note":"tacos and coffee","source":"cli","meta":{}}
```

---

## Why JSON matters

Because now you can:

* graph your life
* build a desktop app later
* analyze patterns
* combine with other tools (like Gravity)

Pulse is not just a CLI.

It is a **structured record of your life**.

---

## The Bigger Picture

Pulse is one part of a system:

```text
Gravity → where your work went
Pulse   → how you lived life 
```

Together:

```text
You stop guessing.
You start seeing.

```


---

## UPDATE: Universal Tag System

Pulse now supports tags on every event.

Tags look like this:

    tag:<name>

They’re saved as structured JSON under:

    meta.tags

So your note stays clean and human,
and the tags quietly keep track of context behind the scenes.

    note = what you meant to say
    tags = the little helpers that remember extra stuff

------------------------------------------------------------
BASIC SHAPE
------------------------------------------------------------

    pulse <event> ["note"] tag:<tag> tag:<tag>

Examples:

    pulse mood "a little off but okay" tag:softday tag:lightfog
    pulse pain "ankle made a weird click earlier" tag:bodynote tag:keepaneye
    pulse note "the spoon felt important for some reason" tag:spoonmoment tag:homevibes
    pulse dream "a frog ran a bank and I didn’t qualify" tag:frogbank tag:nightlogic

------------------------------------------------------------
REGULAR EVENTS WITH TAGS
------------------------------------------------------------

    pulse mood "kind of confused but still functioning" tag:cloudy tag:lowenergy
    pulse energy "running on vibes and a snack" tag:lowpower tag:stillgoing
    pulse note "the window looked different today" tag:windowmoment tag:quietthought
    pulse dream "a frog was running a very small business" tag:frogbank tag:tinyworld

Saved shape:

    {
      "type": "mood",
      "category": "mental",
      "note": "kind of confused but still functioning",
      "source": "cli",
      "meta": {
        "tags": ["cloudy", "lowenergy"]
      }
    }

------------------------------------------------------------
TAG-ONLY EVENTS
------------------------------------------------------------

You can also just leave tags and move on.

    pulse mood tag:cloudy tag:lowenergy
    pulse pain tag:anklething tag:minorissue
    pulse note tag:quickthought tag:latermaybe

Saved shape:

    {
      "type": "mood",
      "category": "mental",
      "source": "cli",
      "meta": {
        "tags": ["cloudy", "lowenergy"]
      }
    }

Good for when you don’t feel like explaining everything.

------------------------------------------------------------
TOGGLE EVENTS WITH TAGS
------------------------------------------------------------

Start and end moments can have their own tags.

    pulse lunch "had soup and sat for a bit" tag:warmmeal tag:slowmoment
    pulse lunch "finished eating, back to it" tag:done tag:reset

    pulse walk "went outside for a short walk" tag:freshair tag:sidequest
    pulse walk "back home again" tag:return tag:smallwin

    pulse work "started working at the computer" tag:focus tag:onscreen
    pulse work "wrapped things up for now" tag:done tag:steppingaway

Each part keeps its own context.

------------------------------------------------------------
IMPORTANT EVENTS
------------------------------------------------------------

Important events are things you might want to remember later.

    pulse important "realized I like working in the terminal" tag:smalljoy tag:glowbox
    pulse important "cleaned up my space and it helped" tag:reset tag:calm
    pulse important "found something where it shouldn’t have been" tag:oddmoment tag:noticed
    pulse important "launched comic2panel — first clean run" tag:milestone tag:build
    pulse important "pushed code and it actually worked" tag:shipit tag:relief

Then:

    pulse year
    pulse year 2026

Shows only important events from that year.

Example:

    Pulse — 2026

    May:
    2026-05-02  realized I like working in the terminal  [smalljoy, glowbox]
    2026-05-02  cleaned up my space and it helped  [reset, calm]
    2026-05-02  found something where it shouldn’t have been  [oddmoment, noticed]
    2026-05-02  launched comic2panel — first clean run  [milestone, build]

    4 important events found.

------------------------------------------------------------
TINY SILLY COMMAND GARDEN
------------------------------------------------------------

    pulse wake "dreamed something strange but kind of nice" tag:softdream tag:earlymorning
    pulse mood "a bit slow but alright" tag:lowkey tag:steady
    pulse lunch "ate noodles and took a break" tag:noodles tag:pause
    pulse note "the house made a sound, probably fine" tag:housemoment tag:shrug
    pulse work "did some work at the computer" tag:focus tag:progress
    pulse important "made a small tool that feels useful" tag:build tag:smallwin
