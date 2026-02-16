# Local database setup

Use a local MongoDB and seed data when the cloud API is unavailable or you want to work offline.

## 1. Use local MongoDB

Copy the local env example and edit `.env` to point at your local MongoDB:

```bash
cp env.local.example .env
# Edit .env and set:
# MONGODB_URI=mongodb://localhost:27017/krm
```

Start MongoDB locally (e.g. `mongod` or Docker: `docker run -p 27017:27017 mongo`).

## 2. Seed admin and proposals

```bash
npm run seed:admin    # creates admin@krm.com / admin123456
npm run seed:proposals # loads backend/data/proposals-dump.json
# Or both:
npm run seed
```

## 3. Run backend and frontend

```bash
# Terminal 1 - backend
npm run dev

# Terminal 2 - frontend (from repo root)
cd ../frontend && npm run dev
```

Log in with **admin@krm.com** / **admin123456**. The proposals list will show the seeded proposals.

## Customising the dump

Edit `backend/data/proposals-dump.json` to add or change proposals. Each object needs:

- `name`, `client`, `project`, `template`
- `rawExcelData`: `fileName`, `sheetName`, `headers` (array), `rows` (array of row arrays)

Then run `npm run seed:proposals` again. To avoid duplicates, clear the collection first or use a fresh local DB.
