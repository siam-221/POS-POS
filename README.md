# Alien Tees ERP — সংশোধিত ফাইল ব্যবহারের নির্দেশনা

## এই কপিতে কী সংশোধন করা হয়েছে

- মূল আপলোড করা ফাইলটি অপরিবর্তিত রাখা হয়েছে; সংশোধিত কপি আলাদা তৈরি করা হয়েছে।
- স্থায়ী random Flask secret ব্যবহার করা হয়েছে; predictable hard-coded secret fallback সরানো হয়েছে।
- নতুন database-এ `ADMIN_PASS` না দেওয়া থাকলে শক্তিশালী one-time Super Admin password তৈরি হয়ে server log-এ দেখাবে। আগের user/password স্বয়ংক্রিয়ভাবে বদলানো হয় না।
- Failed login rate limit যোগ করা হয়েছে: একই username/IP hash-এর ১৫ মিনিটে ৫টি ব্যর্থ চেষ্টার পর সাময়িক block।
- Role database থেকে পুনরায় যাচাই করা, password-এর ন্যূনতম দৈর্ঘ্য এক করা এবং cookie configuration উন্নত করা হয়েছে।
- Browser-এর state-changing request-এ Origin/Referer যাচাই যোগ করা হয়েছে। Origin/Referer ছাড়া সরাসরি API request এখন reject হবে।
- Stock workbook access-এ thread lock-এর পাশাপাশি OS file lock যোগ করা হয়েছে, যাতে একই সংশোধিত কোড চালানো একাধিক process workbook-এ একসঙ্গে পরিবর্তন না করে।
- Stock Entry-তে idempotency request ID যোগ করা হয়েছে। একই pending request retry করলে quantity দ্বিতীয়বার যোগ হওয়ার ঝুঁকি কমে।
- Stock edit/upload/reconciliation, POS stock deduction, Purchase Receive Stock এবং purchase delete-এ workbook rollback ও history consistency protection উন্নত করা হয়েছে।
- Stock quantity parsing, row/code/size validation এবং stock-add-only ব্যবহারকারীর জন্য quantity গোপন রাখার নিয়ম উন্নত করা হয়েছে।
- Stock comment deletion-এর permission configured permission policy অনুযায়ী চালানো হয়েছে।
- HR Attendance-এ date/time validation কঠোর করা, manual overtime hours/minutes এবং OT adjustment audit log যোগ করা হয়েছে। একজন কর্মীর মাসিক report-এ record না থাকা দিনগুলো `No Record` হিসেবে দেখানো হয়, যাতে সেগুলো ভুল করে স্বয়ংক্রিয়ভাবে absent গণ্য না হয়।

## ব্যবহার করার আগে গুরুত্বপূর্ণ ধাপ

1. বর্তমান অ্যাপের **সব worker/process বন্ধ করুন**। পুরোনো worker নতুন file lock ব্যবহার করবে না, তাই পুরোনো worker চালু থাকলে stock workbook নিয়ে দ্বন্দ্ব হতে পারে।
2. বর্তমান `data.db`, `instance/stock_master.xlsx`, `instance/stock_backups/` এবং প্রয়োজনীয় অন্য persistent files-এর backup নিন।
3. ZIP-এর `app.py` অথবা `app_fixed.py`-কে বর্তমান `app.py`-এর **একই folder-এ** রাখুন। পুরোনো `app.py` আলাদা নামে backup করে নতুন ফাইলটির নাম `app.py` করুন। অন্য folder থেকে চালালে নতুন/ফাঁকা database তৈরি হতে পারে, কারণ path Python ফাইলটির অবস্থান থেকে নির্ধারিত হয়।
4. Dependency install করুন:

   ```bash
   python -m pip install -r requirements.txt
   ```

5. Environment variable নির্ধারণ করুন। Production-এ কমপক্ষে ৩২ অক্ষরের random `SECRET_KEY`, উপযুক্ত `ADMIN_USER` এবং HTTPS চালু থাকলে `COOKIE_SECURE=1` ব্যবহার করুন। `ADMIN_PASS` শুধু নতুন database-এ প্রথম Super Admin তৈরি করার জন্য; বিদ্যমান user-এর password এটি বদলায় না।
6. নতুন database-এ `ADMIN_PASS` না দিলে startup log-এ তৈরি হওয়া one-time bootstrap password দেখাবে। সেটি নিরাপদে সংরক্ষণ করে login করার পর বদলে নিন। Server log অন্য কেউ দেখতে পারে এমনভাবে প্রকাশ করবেন না।
7. Reverse proxy ব্যবহার করলে public `Host` header সঠিকভাবে forward হতে হবে, নইলে Origin/Referer check legitimate request-ও block করতে পারে। Browser-এর বাইরে থেকে API call করলে Origin/Referer না পাঠালে request reject হবে।
8. প্রথমে test copy-তে login, Guest access, Stock Entry, stock edit, POS sale, Purchase Receive Stock, purchase delete, backup/restore, HR attendance, overtime ও payroll পরীক্ষা করে তারপর live কাজে ব্যবহার করুন।

## এই পরিবেশে কী পরীক্ষা করা হয়েছে

- Python syntax compilation সফল।
- Python AST parsing সফল; ১৪৮টি Flask route decorator পাওয়া গেছে এবং একই method/path duplicate পাওয়া যায়নি।
- `HTML` ও `LOGIN_HTML`-এর মধ্যে থাকা চারটি JavaScript block `node --check` দিয়ে সফলভাবে পরীক্ষা করা হয়েছে।
- **Live Flask smoke test চালানো যায়নি**, কারণ এই execution environment-এ Flask/Werkzeug ইনস্টল ছিল না এবং package index-এ network access পাওয়া যায়নি। তাই deployment server-এ integration test বাধ্যতামূলক। Syntax check সফল হওয়া runtime বা ব্যবসায়িক হিসাবের সম্পূর্ণ নিশ্চয়তা নয়।

## এখনও আলাদা পরিকল্পনা ও পরীক্ষার প্রয়োজন

- Excel এবং SQLite এখনও আলাদা data store। Rollback protection উন্নত করা হয়েছে, তবে এটি একক database transaction-এর সমান নয়। বড় multi-worker deployment-এর দীর্ঘমেয়াদি ভালো সমাধান হলো database-based stock ledger।
- Barcode scanner workflow, পূর্ণ customer payment collection ledger, সব sales channel অন্তর্ভুক্ত audited accounting, এবং automated end-to-end test suite এই patch-এ নতুন করে সম্পূর্ণ implement করা হয়নি; এগুলোর জন্য আলাদা business rules ও integration tests দরকার।
