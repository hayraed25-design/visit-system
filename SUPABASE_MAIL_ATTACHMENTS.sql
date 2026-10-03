

create table if not exists mail_attachments (
  id bigint generated always as identity primary key,
  mail_id bigint not null references mail(id) on delete cascade,
  file_name text not null,
  storage_path text not null unique,
  file_size bigint,
  mime_type text,
  created_at timestamptz not null default now()
);

alter table mail_attachments enable row level security;

drop policy if exists "mail_attachments_select" on mail_attachments;
drop policy if exists "mail_attachments_insert" on mail_attachments;
drop policy if exists "mail_attachments_delete" on mail_attachments;

create policy "mail_attachments_select"
on mail_attachments
for select
to anon
using (true);

create policy "mail_attachments_insert"
on mail_attachments
for insert
to anon
with check (true);

create policy "mail_attachments_delete"
on mail_attachments
for delete
to anon
using (true);

insert into storage.buckets (id, name, public)
values ('mail-attachments', 'mail-attachments', false)
on conflict (id) do update set public = false;

drop policy if exists "mail_attachments_storage_insert" on storage.objects;
drop policy if exists "mail_attachments_storage_select" on storage.objects;
drop policy if exists "mail_attachments_storage_delete" on storage.objects;

create policy "mail_attachments_storage_insert"
on storage.objects
for insert
to anon
with check (bucket_id = 'mail-attachments');

create policy "mail_attachments_storage_select"
on storage.objects
for select
to anon
using (bucket_id = 'mail-attachments');

create policy "mail_attachments_storage_delete"
on storage.objects
for delete
to anon
using (bucket_id = 'mail-attachments');


-- حذف البريد من التطبيق (مع حذف سجلات المرفقات المرتبطة بسبب ON DELETE CASCADE)
drop policy if exists "mail_delete" on mail;
create policy "mail_delete"
on mail
for delete
to anon
using (true);
