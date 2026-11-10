# Put this site on GitHub, then Render

Do not upload this zip as a single file inside a GitHub repository. Render builds the source files. A zip sitting in the repo will not run.

## 1. GitHub

1. On a computer, unzip this folder.
2. Create a new empty GitHub repository. Do not add a README, so the repo stays empty.
3. In the unzipped folder, run:

   git init
   git add .
   git commit -m "Fontenot engineering report"
   git branch -M main
   git remote add origin YOUR_GITHUB_REPO_URL
   git push -u origin main

## 2. Render

1. In Render, choose New, then Blueprint.
2. Select that GitHub repository.
3. Apply the Blueprint. It creates the web service and a Postgres database.
4. The first deploy installs dependencies, builds the site, and creates the tables.

Render charges a small monthly amount for the web service and the database. There is no lasting free database on Render.

## 3. After it is up

Open the Render URL. Use Admin to create the admin email and password. That login is new. It is not the Grok login from the other site.

Then upload the Quore files again and tap Run engine. Saved weeks from msyps-engineering.grok.me are not inside this zip. They stay in that site's database. This copy starts from the sample week bundled with the code, then keeps whatever is uploaded on Render.

## What Render needs

- Node 22
- Build: npm install && npm run build
- Start: node .output/server/index.mjs
- DATABASE_URL from the Render Postgres database
- NITRO_PRESET=render_com
- BETTER_AUTH_SECRET, which the Blueprint generates

The public report does not require sign-in. Sign-in is only for uploading a new week.
