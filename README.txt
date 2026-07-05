Journal Publishing System v2.0 REAL
===================================

This is a real working Python/Tkinter app. It does not automatically call ChatGPT.
It prepares prompt files, lets you process them through ChatGPT manually, and merges
ChatGPT's JSON responses into CSV databases.

INSTALL
-------
1. Unzip this folder.
2. Double-click Install_Requirements.bat.
3. Wait until it finishes.

RUN
---
1. Double-click Start_Journal_Publishing_System.bat.
2. Click Select Journal .docx.
3. Choose your Word journal.
4. Set Words per chunk. Recommended: 1500 to 2000.
5. Click Create Hierarchical Prompt Files.

PROCESS CHUNKS THROUGH CHATGPT
------------------------------
1. Click Open prompts_to_paste.
2. Open Prompt_0001_....txt.
3. Copy the entire prompt.
4. Paste it into ChatGPT.
5. ChatGPT should return JSON only.
6. Save that JSON response as Chunk_0001.json in the ai_json_results folder.
7. Repeat for Prompt_0002, Prompt_0003, etc.

MERGE RESULTS
-------------
1. After you have saved JSON result files in ai_json_results, return to the app.
2. Click Merge JSON Results into CSVs.
3. Click Open output.

OUTPUT FILES
------------
The output folder will contain:
- Master_Database.csv
- Journal_Entries.csv
- Essays.csv
- Paragraphs.csv
- One_Liners.csv
- Principles.csv
- Stories.csv
- Definitions.csv
- Book_Ideas.csv
- Questions.csv
- Personal_Testimonies.csv
- Book_View folder, with one CSV per book/project tag.

IMPORTANT
---------
The prompts ask ChatGPT to preserve hierarchy:
Journal entry -> Essay -> Paragraph -> Quote/Principle/Story/etc.

Each extracted row includes:
item_id, parent_id, journal_entry_id, essay_id, paragraph_id, sequence.

This means a one-liner can be linked back to its paragraph, essay, and journal entry.
