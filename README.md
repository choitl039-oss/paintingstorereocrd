# Jennifer Store Keeping

A static inventory website for paintings that can be hosted on GitHub Pages.

What it does:
- Add a painting with drag-and-drop image upload, name, price, and quantity.
- Let staff record how many were sold.
- Automatically reduce stock after each sale.
- Keep a live sales log.
- Export an Excel workbook with Summary, Inventory, and Sales sheets.

## Files to upload to GitHub

Keep these files in your repository root:
- index.html
- styles.css
- app.js
- supabase-config.js
- README.md

## Supabase setup

This website needs Supabase because GitHub Pages only hosts static files. Supabase stores shared stock and sales records online so your staff can see the same data.

1. Create a Supabase project at https://supabase.com/
2. In Supabase SQL editor, create the required tables and policies (schema below).
3. Storage bucket is optional in this version (images are kept for display only in records).
4. Open `supabase-config.js` and replace all `REPLACE_WITH_...` values.
5. Change `adminPin` in `supabase-config.js` to your own PIN.

Example:

```js
export const supabaseConfig = {
  url: 'https://your-project-ref.supabase.co',
  anonKey: 'your-anon-public-key',
  storageBucket: 'painting-images',
};

export const adminPin = '5678';
```

Note: Excel export includes name, quantity, price, sold units, and totals. It does not include image data.

SQL schema (run once):

```sql
create table if not exists paintings (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  price numeric not null default 0,
  quantity integer not null default 0,
  total_sold integer not null default 0,
  image_url text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists sales (
  id uuid primary key default gen_random_uuid(),
  painting_id uuid references paintings(id) on delete set null,
  item_name text not null,
  unit_price numeric not null default 0,
  quantity_sold integer not null default 0,
  total_amount numeric not null default 0,
  staff_name text,
  sold_at timestamptz not null default now()
);
```

For quick testing, enable row-level security and allow public access (tighten later):

```sql
alter table paintings enable row level security;
alter table sales enable row level security;

create policy if not exists paintings_public_all on paintings
for all using (true) with check (true);

create policy if not exists sales_public_all on sales
for all using (true) with check (true);
```

## GitHub Pages setup

1. Create a new GitHub repository.
2. Upload all project files into the repository root.
3. In GitHub, open Settings > Pages.
4. Under Build and deployment, choose Deploy from a branch.
5. Select the main branch and the /(root) folder.
6. Save the settings.
7. GitHub will give you a website link.

## How the app works

Admin view:
- Unlock with the admin PIN.
- Drag and drop a painting photo to add stock faster.
- Update the name, price, and quantity of any existing painting.
- Export the Excel report.

Staff view:
- See the painting cards.
- Enter staff name if needed.
- Press Record sale and enter how many were sold.

Reports:
- The top summary shows total cash received, total units sold, inventory value, and number of paintings.
- Export Excel report downloads a workbook with all current records.

## Important note about security

The current admin PIN is a simple browser-side lock. It is useful for basic separation between owner and staff views, but it is not strong security.

If you want real staff accounts and proper permissions, the next step is to add Firebase Authentication and Firestore security rules.

## Local testing

You can open index.html directly in a browser, but Firebase features work more reliably when served from a simple web server.

If you later install Python on this computer, you can run:

```powershell
python -m http.server 8000
```

Then open http://localhost:8000
