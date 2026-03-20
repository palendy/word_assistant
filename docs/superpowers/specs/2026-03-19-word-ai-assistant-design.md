# MS Word AI Assistant Add-in Design

## Overview

On-prem MS Word Add-in that provides a chat-based AI assistant in a side panel (TaskPane). Users open a Word template, paste raw data (e.g. quarterly financial results) into the chat, and the AI automatically fills in the document according to the template structure.

## Requirements

- **Platform**: Windows desktop Word (develop/test on Mac)
- **Distribution**: Sideload via manifest.xml (no marketplace)
- **AI Backend**: OpenAI-compatible API (vLLM, Ollama, etc.) with configurable URL
- **Authentication**: None — install and configure URL to use
- **Core Feature**: Template-based document generation from raw text data via chat

## Architecture

```
MS Word (Desktop)
├── Document Area (template.docx with tables, text, formatting)
└── TaskPane (side panel)
    ├── Chat UI (user inputs raw data, AI responds)
    └── Settings (AI server URL, model name)
        │
        │  fetch (OpenAI-compatible API)
        ▼
    AI Server (on-prem, vLLM/Ollama/etc.)
```

## Tech Stack

- **Framework**: Office.js Web Add-in
- **Frontend**: React + TypeScript
- **Bundler**: Webpack
- **Styling**: CSS (lightweight, no framework)
- **API Protocol**: OpenAI Chat Completions (`/v1/chat/completions`)

## Project Structure

```
09_word_extension/
├── manifest.xml              # Office Add-in manifest (sideload)
├── package.json
├── tsconfig.json
├── webpack.config.js
├── src/
│   ├── taskpane/
│   │   ├── index.html        # TaskPane entry point
│   │   ├── index.tsx         # React mount
│   │   ├── App.tsx           # Main app (chat/settings tabs)
│   │   ├── components/
│   │   │   ├── ChatPanel.tsx     # Chat UI
│   │   │   ├── MessageBubble.tsx # Message bubbles
│   │   │   ├── SettingsPanel.tsx  # AI URL/model settings
│   │   │   └── StatusBar.tsx     # Connection status
│   │   ├── services/
│   │   │   ├── aiClient.ts       # OpenAI-compatible API calls
│   │   │   ├── wordDocument.ts   # Office.js document read/write
│   │   │   └── settings.ts       # Settings persistence (localStorage)
│   │   └── styles/
│   │       └── app.css
│   └── commands/
│       └── commands.ts       # Ribbon button commands
└── assets/
    └── icon-*.png            # Icons
```

## Core Workflow

### Step 1: Read Document Structure (`wordDocument.ts`)
- Use Office.js API to read current document body: paragraphs, tables, content controls
- Convert to structured text representation for AI context

### Step 2: Build AI Prompt (`aiClient.ts`)
- System prompt: instructs AI to fill template structure with provided data
- User message: document structure + raw data from chat input
- Call OpenAI-compatible `/v1/chat/completions` endpoint

### Step 3: Parse Response & Modify Document (`wordDocument.ts`)
- Parse AI response (structured format: JSON or marked sections)
- Use Office.js API to insert/replace text in paragraphs, table cells, etc.

## AI Integration

### API Configuration
- **URL**: User-configurable (e.g. `http://server:8000/v1/chat/completions`)
- **Model**: User-configurable model name
- **Storage**: `localStorage` for persistence

### Prompt Strategy
- Send document structure as context so AI understands the template layout
- AI responds with structured instructions for what text goes where
- Add-in parses and applies changes via Office.js

## Deployment

### Sideload Method
1. Host add-in files on internal web server (or localhost for dev)
2. Configure `manifest.xml` with server URL
3. Sideload via:
   - Network share folder (group policy)
   - Manual sideload in Word (dev/testing)

## Development & Testing

| Phase | Method |
|-------|--------|
| Chat UI | Browser dev server (localhost) |
| Office.js integration | Sideload on Mac Word |
| AI API | Local Ollama or mock server |
| End-to-end | Mac Word + local AI server |
