A small internal tool I built to eliminate manual data entry at work.

## The problem

Our team was required by a government mandate to conduct surveys on paper, then manually transcribe every response into LimeSurvey to generate statistics. It was tedious, error-prone, and a waste of time.

## The solution

I built a web-based version of the survey that staff can fill out digitally. Responses are collected via a serverless API and stored directly — no paper, no manual re-entry, same stats at the end.

## Tech stack

- **Frontend** — Vanilla HTML, CSS, JavaScript
- **Backend** — Serverless functions (Node.js) via Vercel
- **Deployment** — Vercel

## Project structure

```
ankiety/
├── api/        # Serverless API endpoints
├── public/     # Frontend (HTML, CSS, JS)
├── vercel.json # Routing & deployment config
└── package.json
```

## Running locally

```bash
git clone https://github.com/strxless/ankiety.git
cd ankiety
npm install
npx vercel dev
```
