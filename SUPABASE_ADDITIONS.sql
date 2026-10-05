-- إضافات نظام الزيارات: سجل الاتصالات + المواعيد
-- يمكن تشغيل هذا الملف في Supabase SQL Editor.

/* =========================
   سجل الاتصالات
========================= */

create table if not exists calls (
  id bigint generated always as identity primary key,
  call_no text,
  caller_name text not null,
  company_name text not null,
  phone text not null,
  subject text not null,
  requested_employee text not null,
  status text not null default 'تحويل مكالمة',
  created_at timestamptz not null default now()
);

alter table calls add column if not exists call_no text;
alter table calls add column if not exists caller_name text;
alter table calls add column if not exists company_name text;
alter table calls add column if not exists phone text;
alter table calls add column if not exists subject text;
alter table calls add column if not exists requested_employee text;
alter table calls add column if not exists status text default 'تحويل مكالمة';
alter table calls add column if not exists created_at timestamptz default now();

alter table calls enable row level security;

drop policy if exists "calls_select" on calls;
drop policy if exists "calls_insert" on calls;
drop policy if exists "calls_update" on calls;

create policy "calls_select"
on calls
for select
to anon
using (true);

create policy "calls_insert"
on calls
for insert
to anon
with check (true);

create policy "calls_update"
on calls
for update
to anon
using (true)
with check (true);

/* =========================
   المواعيد
========================= */

create table if not exists appointments (
  id bigint generated always as identity primary key,
  appointment_no text,
  person_name text not null,
  day_name text not null,
  entry_date date not null,
  appointment_date date not null,
  appointment_time time not null,
  host_name text not null,
  status text not null default 'قيد الانتظار',
  postponement_type text,
  postponed_date date,
  created_at timestamptz not null default now()
);

alter table appointments add column if not exists appointment_no text;
alter table appointments add column if not exists person_name text;
alter table appointments add column if not exists day_name text;
alter table appointments add column if not exists entry_date date;
alter table appointments add column if not exists appointment_date date;
alter table appointments add column if not exists appointment_time time;
alter table appointments add column if not exists host_name text;
alter table appointments add column if not exists status text default 'قيد الانتظار';
alter table appointments alter column status set default 'قيد الانتظار';
alter table appointments add column if not exists postponement_type text;
alter table appointments add column if not exists postponed_date date;
alter table appointments add column if not exists created_at timestamptz default now();

alter table appointments enable row level security;

drop policy if exists "appointments_select" on appointments;
drop policy if exists "appointments_insert" on appointments;
drop policy if exists "appointments_update" on appointments;

create policy "appointments_select"
on appointments
for select
to anon
using (true);

create policy "appointments_insert"
on appointments
for insert
to anon
with check (true);

create policy "appointments_update"
on appointments
for update
to anon
using (true)
with check (true);

-- ملاحظة:
-- رقم الاتصال C-000001 ورقم الموعد A-000001 يتم توليدهما من التطبيق
-- مباشرة بعد إنشاء السجل، لذلك لا نحتاج إلى Trigger إضافي.

/* =========================
   إصلاح حالات المواعيد
   السماح بالحالات الجديدة المطلوبة
========================= */
alter table appointments drop constraint if exists appointments_status_check;

-- تحويل القيم القديمة إلى القيم الجديدة قبل إضافة القيد
update appointments set status = 'قيد الانتظار'
where status is null or trim(status) = '' or status = 'انتظار';

update appointments set status = 'تم الموعد'
where status = 'تم';

update appointments set status = 'تأجيل الموعد'
where status = 'تأجيل';

update appointments set status = 'إلغاء الموعد'
where status = 'إلغاء';

-- أي قيمة قديمة أخرى غير معروفة تصبح الحالة الافتراضية
update appointments set status = 'قيد الانتظار'
where status not in ('قيد الانتظار', 'تم الموعد', 'تأجيل الموعد', 'إلغاء الموعد');

alter table appointments add constraint appointments_status_check
check (status in ('قيد الانتظار', 'تم الموعد', 'تأجيل الموعد', 'إلغاء الموعد'));

alter table appointments alter column status set default 'قيد الانتظار';

/* =========================
   البريد الإلكتروني للشركات
========================= */

alter table companies
add column if not exists email text;
