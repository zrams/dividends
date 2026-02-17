# Family Dinner Planner (React + Express + MongoDB)

Initial full-stack project structure for a family dinner planning app.

## Features included

- React frontend (Vite)
- Node.js + Express backend
- MongoDB with Mongoose models
- JWT authentication
- Role-based access:
  - `admin` (mom): add dinner ideas, view all weekly submissions
  - `member`: submit 1-2 dinner choices for a week
- Schemas:
  - `User`: `name`, `email`, `password`, `role`
  - `Dinner`: `name`, `description`
  - `WeeklySubmission`: `userId`, `weekStartDate`, `choices[]`

## Project structure

```text
family-dinner-planner/
  backend/
    src/
      config/db.js
      middleware/auth.js
      models/
      routes/
      server.js
    .env.example
    package.json
  frontend/
    src/
      App.jsx
      api.js
      main.jsx
      styles.css
    .env.example
    index.html
    package.json
    vite.config.js
  package.json
  .gitignore
```

## 1) Install dependencies

From the `family-dinner-planner` directory:

```bash
npm install
npm install --prefix backend
npm install --prefix frontend
```

> `npm install` at root installs `concurrently` for running both apps together.

## 2) Configure environment variables

### Backend

```bash
cp backend/.env.example backend/.env
```

Edit `backend/.env`:

- `MONGO_URI` (local MongoDB or Atlas connection string)
- `JWT_SECRET` (any secure string for local dev)
- `PORT` (default `5000`)
- `CLIENT_URL` (default `http://localhost:5173`)

### Frontend

```bash
cp frontend/.env.example frontend/.env
```

Default value points to local backend:

- `VITE_API_URL=http://localhost:5000/api`

## 3) Run locally with npm

From `family-dinner-planner`:

```bash
npm run dev
```

This starts:

- Backend: `http://localhost:5000`
- Frontend: `http://localhost:5173`

## API routes (initial)

### Auth

- `POST /api/auth/register`
- `POST /api/auth/login`

### Dinner ideas

- `GET /api/dinners` (authenticated user)
- `POST /api/dinners` (admin only)

### Weekly submissions

- `POST /api/submissions` (member only, 1-2 choices)
- `GET /api/submissions/mine` (member only)
- `GET /api/submissions` (admin only)

## Notes

- Register mom as an `admin` account.
- Members submit top 1-2 dinner picks per week.
- Admin can add 50-100 dinner ideas over time and review all submissions.
