# Generic Syllabus-Driven Markdown Presentation Skill

## Purpose
This skill is for creating **lecture presentations in markdown** that are later converted into slide decks by a markdown-to-presentation pipeline.

It is **generic** across subjects.
It must **adapt automatically to the subject, course naming, lecture numbering, topic order, terminology, case selection, and file naming** based on the **provided syllabus** and any previously supplied lecture markdown files.

This skill must never hardcode one subject area into the workflow.
The syllabus is the source of truth.

---

## Core principle
**The subject comes from the syllabus, not from the skill.**

The skill defines:
- the markdown presentation format,
- strict slide-writing rules,
- image and asset rules,
- merging and packaging workflow,
- quality checks.

The syllabus defines:
- the subject name,
- the course sequence,
- topic order,
- lecture scope,
- terminology,
- expected methods,
- academic emphasis,
- appropriate examples and case studies.

If the syllabus says the course is Big Data, Operating Systems, Marketing Analytics, Databases, or Ethics, the skill must adapt.

---

## Required inputs
Before writing slides, gather these inputs whenever available:

1. **Syllabus**
   - Primary source of truth.
   - Use it to identify course name, module order, topic order, learning goals, and terminology.

2. **Previous lecture markdown files**
   - Use them to infer writing style, depth, lecture continuity, layout vocabulary, markdown syntax, and pacing.

3. **User instructions from the current conversation**
   - These override defaults.
   - Treat all explicit user constraints as binding.

4. **Source readings, publications, case materials, or uploaded references**
   - Use them to support the lecture academically.
   - If the lecture depends on papers or reports, download them first and prepare them for figure extraction and reading.

---

## Source Paper Workflow
When the lecture uses external papers, reports, or long web sources, follow this order before slide writing:

1. Identify the final source list from the syllabus and user request.
2. Download the papers and web sources locally.
3. If any requested paper is missing, download it before moving on.
4. Convert the papers to readable markdown and extract embedded figures or page images with the `PicAxe-OCR` pipeline.
5. Prefer the batch helper `bash tools/prepare_lecture_sources.sh <lecture_dir>` so download and conversion happen together, with a small batch size when several PDFs are present. Always pass the actual lecture folder for the current deck; do not rely on a hardcoded default lecture path.
6. Keep the source list data-driven: store the lecture-specific URLs in `<lecture_dir>/sources.tsv` and let the downloader read that manifest instead of hardcoding topic URLs into the script.
7. Keep the conversion outputs in a stable folder structure so later slide batches can reuse them. The helper should write a sibling `<lecture_dir>/literature/` folder next to the lecture markdown, with mirrored `papers/`, `web/`, `extracted_images/`, `tables/`, and a consolidated bundle such as `<lecture_dir>/literature/picaxe-out/` with per-batch subfolders.
8. If the conversion is slow on the current machine, generate one terminal command that performs the download and conversion end-to-end and tell the user to run it offline before resuming lecture generation.

Use the converted markdown and extracted figures as the working source set for planning, slide writing, and visual asset selection. Do not start slide writing from an incomplete source set if the missing papers can still be downloaded.

---

## Required adaptation from syllabus
For every new lecture, the skill must derive the following from the syllabus and context:

### 1. Course subject
Identify the course or module title from the syllabus.
Use this to determine the lecture’s conceptual domain.

### 2. Lecture position in the course
Determine where the requested lecture sits in the sequence.
Infer:
- what was likely covered earlier,
- what should be introduced now,
- what should be deferred to later lectures.

### 3. Correct scope
Do not teach material too early.
Do not repeat earlier material unless a short bridge is needed.
Do not introduce later-course techniques in depth unless the syllabus supports it.

### 4. Naming prefix
The file prefix must be derived from the syllabus course name, not hardcoded.

#### Naming algorithm
Given a course title from the syllabus:
1. Convert to lowercase.
2. Remove punctuation.
3. Remove spaces.
4. Remove diacritics when needed.
5. Keep it concise.
6. Use a meaningful course-level prefix.

Examples:
- `Big Data` -> `bigdata`
- `Operating Systems` -> `operatingsystems`
- `Machine Learning` -> `machinelearning`
- `Digital Marketing Analytics` -> `digitalmarketinganalytics`

Then append lecture number:
- `bigdata03`
- `operatingsystems05`
- `machinelearning02`

If the user explicitly requests a different prefix, use the user’s requested naming.

### 5. Case and example selection
Examples must match the syllabus topic and lecture position.
Use real-life examples, but only those that fit the current lecture.

### 6. Terminology
Terminology must follow the syllabus and previous lecture style.
Do not switch fields or vocabulary casually.

---

## Output expectations
The skill must produce:

1. **Markdown lecture files**
2. **PNG image assets**
3. **Merged cumulative markdown after each step**
4. **A zip package containing**:
   - the merged markdown file,
   - all referenced image assets,
   - optionally partial continuation markdown files if requested.
5. **A presentation**:
   - pptx,
   - pdf,
   - generat with slidecanon.py based on 'README.md'.

---

## Strict markdown slide format
The markdown must follow the existing slide syntax used by the user’s presentation pipeline.

Each slide begins with:

```md
#slide <number>
```

Then one or more structured sections such as:

```md
## layout
bullet
## slide title
Example title
## subtitle
Example subtitle
## bullets
- Bullet 1
- Bullet 2
## image
![Alt text](file.png)
## teleprompter:
Narration text here.
```

### Allowed fields
Use only fields already supported by the user’s existing markdown presentation format, such as:
- `## layout`
- `## slide title`
- `## subtitle`
- `## bullets`
- `## term`
- `## definition`
- `## story`
- `## image`
- `## left title`
- `## left bullets`
- `## right title`
- `## right bullets`
- `## teleprompter:`

Do not invent new markdown fields unless the existing source files clearly use them.

---

## Allowed layouts
Default to layouts already seen in the existing lecture files.

Common safe layouts:
- `title`
- `bullet`
- `definition`
- `story`
- `comparison`
- `bullet + image`

### Important rule
Prefer:
- `bullet + image`

Avoid:
- `bullet + diagram`
- `bullet + photo`

Reason:
All diagrams, charts, plots, equations, schematics, and illustrations should be handled as **image assets** placed via markdown image references.
If assets are generated using scripts, then also save these scripts.

---

## Mandatory image placement rules
### 1. All visuals are separate PNG assets
Every diagram, chart, plot, equation, schematic, or structured illustration must be created as a separate PNG asset.

### 1a. Chart generation method
For charts and data visualizations, prefer an HTML + CSS + JavaScript workflow.

Use this sequence:
1. build the chart or graphic as a self-contained HTML document,
2. style it with CSS for slide-scale readability,
3. generate data-driven marks, labels, and annotations with JavaScript,
4. use a JavaScript charting or visualization library when it improves clarity, such as D3, Plotly, Vega-Lite, ECharts, or Chart.js,
5. render the final result to a transparent PNG for slide use,
6. save the HTML/CSS/JS source next to the exported image.

The PNG remains the slide asset, but the source of record for chart construction should be the HTML/CSS/JS implementation.

### 1b. Glyphs in generated graphics
Text rendered into PNG is not repaired by PowerPoint font substitution. Before rendering, check that the **actual font file used by the renderer** supports every character in the labels, especially Polish letters, Greek letters, arrows, mathematical operators, superscripts and subscripts. A successful render call does not prove this: some renderers silently draw an empty square for a missing glyph.

For formulas, do not assume Unicode superscript digits are available in the chosen font. Prefer a math-capable font with verified glyph coverage, or draw the exponent separately using ordinary digits at a smaller size and raised position. Keep this construction in the image source script so regeneration preserves it. Do not replace mathematical notation with ambiguous plain text.

After export, inspect a rendered sample of each **distinct font or notation style** at slide scale, including at least one formula containing the uncommon characters. Reject and regenerate any asset with squares, replacement characters, missing marks, clipped exponents or unreadable symbols. When many assets share one renderer, check all label strings for unsupported characters programmatically; visual inspection of every slide is unnecessary.

### 2. Transparent background required
Every image asset must have:
- transparent background,
- no slide-colored background,
- no decorative frame,
- no border unless analytically necessary.

### 3. No baked-in slide furniture
An image asset must not contain:
- slide title,
- subtitle,
- speaker notes,
- captions intended for the slide body,
- page furniture,
- branding,
- extra explanatory notes.

The image is a **slide component**, not a full slide.

### 4. Correct aspect ratio for layout
If layout is `bullet + image`, the image must be designed for a **side-column placement**.
That means:
- portrait composition matching the current template photo placeholder,
- target aspect ratio **3:4** (width:height),
- in `lecture_template.pptx`, layout `Title, Bullets & Photo` uses a picture placeholder sized **5892913 x 7857216**, which is exactly **0.7500**,
- visual balance that leaves margins inside the image itself.

If the layout is `title + image` / `Title & Photo`, the image must be designed for the current wide photo placeholder.
That means:
- landscape composition,
- target aspect ratio **3:2** (width:height),
- in `lecture_template.pptx`, layout `Title & Photo` uses a picture placeholder sized **10392886 x 6928593**, which is exactly **1.5000**,
- no unnecessary empty space,

Do not treat these ratios as vague suggestions. They should guide the asset canvas itself so the inserted image fits the template cleanly with minimal cropping.

### 4a. Source paper figures
If a paper already contains a suitable figure, chart, table, or diagram:
- prefer the source figure over a redrawn replacement,
- crop or extract the figure so it stays source traceable,
- if the figure does not match the slide ratio, add transparent padding instead of distorting or re-cropping the content,
- keep labels large enough to survive slide-scale viewing.

Use `presentation-graphics` for this asset work when figures need cleaning, padding, or source-aware composition.

### 5. Image markdown syntax
Images must be inserted like this:

```md
## image
![Descriptive alt text](courseprefix3_example_asset.png)
```

Do not use image prompt text inside the markdown.
Do not leave generator instructions in the presentation file.

---

## Content rules for slides
### 1. Student-facing only
Slides must contain only content intended for the audience.

Never include teacher-facing planning notes such as:
- why this lecture comes now,
- what changed from the previous lecture,
- what students will see next,
- why this case is pedagogically useful,
- this helps students understand,
- lecturer workflow notes,
- internal planning remarks.

These must not appear on the slide.
These should also generally not appear in the teleprompter.

### 2. Teleprompter rules
Teleprompter text must explain the slide content directly.
It may elaborate, interpret, connect, or deepen the slide.
It must not read like private planning notes.
It should be long (at least 70 words, with 140-200 as a goal)

Allowed in teleprompter:
- conceptual explanation,
- analytical interpretation,
- narrative transitions rooted in the subject,
- clarification of examples,
- warnings and distinctions relevant to the material.

Not allowed in teleprompter:
- “this slide is here because...”
- “students should now see...”
- “in the next part I will...”
- “this is pedagogically useful because...”
- any sentence that sounds like lecture-design commentary instead of subject explanation.

### 3. No meta slides unless the user explicitly wants them
Avoid visible slides such as:
- roadmap,
- lecture plan,
- what changed,
- recap of author workflow,
- transition-only slides,
- slide about why the lecture exists.

If a structural bridge is needed, express it through real academic content.

Example:
Instead of:
- `Why this lecture comes now`

Use:
- `From study design to data exploration`
- `Before modeling, the data must be inspected`
- `Shared state changes program correctness`
- `Intervals can improve interpretation`

### 4. Every slide must do real work
Each visible slide must do at least one of the following:
- define a concept,
- explain a distinction,
- show a method,
- compare alternatives,
- analyze a case,
- illustrate a process,
- warn against a common error,
- show analytical evidence,
- connect theory to a real context.

If a slide does none of these, it should be rewritten or removed.

---

## Writing rules for bullets
Bullets must describe:
- concepts,
- observations,
- analytical logic,
- method behavior,
- case properties,
- domain consequences.

Bullets must **not** describe:
- what the lecturer is doing,
- what the lecture will do next,
- what students are expected to feel,
- why the example was chosen for teaching purposes.

### Good bullet examples
- Shared writes create risk of lost updates.
- Heavy tails make averages less representative.
- Quantile bins preserve balanced group sizes.
- Sensor drift can resemble real environmental change.

### Bad bullet examples
- We now move to the next topic.
- This case helps students see the idea.
- The lecture now shifts toward a new concept.
- This is pedagogically useful because it is simple.

---

## Academic quality rules
### 1. University standard
The lecture must feel like a real university lecture, not a blog post and not a product demo.

### 2. Strong conceptual backbone
Every lecture should connect:
- concept,
- method,
- interpretation,
- example,
- caution.

### 3. Publications and sources
When the user wants strong published grounding, the lecture should be based on credible academic or high-quality sources.
The lecture should not become citation clutter, but its conceptual backbone should be academically defensible.

### 4. Appropriate difficulty
Use the syllabus and previous lectures to keep difficulty at the correct level.
Do not oversimplify advanced topics.
Do not overload early lectures with later technical depth.
Make sure there are NO REPETITIONS. 
Do not make the same slide twice, but put in different words. Better fewer slides than repeat

---

## Case study rules
Case studies are strongly encouraged when they help the lecture become concrete.

### Choosing cases
Use cases that satisfy these conditions:
- fit the syllabus topic,
- fit the lecture’s position in the course,
- support multiple slides or methods,
- are analytically rich,
- do not feel random.

### Case continuity
If previous lectures established a running case, continue it when possible.
This improves coherence across lectures.

### Multiple examples
If the user wants more real-life examples, use several distinct cases, but keep them controlled.
Do not make the lecture feel like a list of unrelated anecdotes.

---

## Subject adaptation rules
This skill must work for any subject.

### Subject-specific adaptation examples
#### If syllabus is about Big Data
Use:
- data exploration,
- variable types,
- distributed settings,
- churn, web analytics, IoT, public health, sensor data, business intelligence, etc., depending on lecture scope.

#### If syllabus is about Operating Systems
Use:
- processes, threads, scheduling, memory, synchronization, filesystems, I/O, security, etc., depending on lecture scope.

#### If syllabus is about Databases
Use:
- schema, transactions, normalization, indexing, queries, distributed storage, etc., depending on lecture scope.

#### If syllabus is about Marketing Analytics
Use:
- segmentation, channels, attribution, funnel analysis, customer value, campaign effects, etc., depending on lecture scope.

### General rule
The skill must not contain one fixed discipline.
It must infer the discipline from the syllabus and use fitting concepts, examples, and naming.

---

## Workflow for multi-step lecture creation
### After each batch of slides
If the lecture is created in parts, after each step:
1. append the new slides to the cumulative markdown,
2. keep numbering consistent,
3. keep asset references valid,
4. create or update a merged markdown file,
5. package the merged markdown with all referenced assets in a zip.

### Merging rule
At every stage there should be a current master file, for example:
- `bigdata3_first20.md`
- later merged into `bigdata3_first40.md`
- later merged into final `bigdata3.md` or another user-preferred final name.

The same pattern should generalize to any course prefix derived from the syllabus.

---

## Packaging rules
When the user asks for downloadable output, provide:

### Minimum package contents
- merged markdown file,
- all PNG assets referenced by that markdown file.

### File naming
Use the syllabus-derived course prefix and lecture number.
Examples:
- `bigdata3_first40_assets.zip`
- `operatingsystems6_assets.zip`
- `machinelearning2_assets.zip`

If the user requests another naming convention, follow the user.

---

## Strict quality checklist before delivery
Before delivering a markdown presentation batch, verify all of the following:

### Subject and scope
- [ ] The subject matches the syllabus.
- [ ] The lecture topic matches the correct place in the syllabus.
- [ ] The lecture does not introduce advanced material too early.
- [ ] Examples fit the subject and lecture scope.

### Slide content
- [ ] No teacher-facing notes appear on slides.
- [ ] No teacher-facing notes appear in teleprompter text.
- [ ] Every slide does real academic work.
- [ ] Titles are concept-facing, not planning-facing.
- [ ] Bullets describe content, not lecture workflow.

### Layout
- [ ] Only supported layouts are used.
- [ ] `bullet + image` is used instead of `bullet + diagram`.
- [ ] Image references use markdown image syntax.

### Images
- [ ] Every chart/diagram/equation is a PNG asset.
- [ ] All assets have transparent background.
- [ ] No asset contains slide title or notes.
- [ ] Right-column images are composed for side placement.
- [ ] Full-slide images are composed for full-slide placement.
- [ ] Arrows visibly look like arrows.
- [ ] Elements are aligned correctly.
- [ ] Elements do not overlap unless deliberately designed to do so.
- [ ] Every graphic font supports the characters actually used; mathematical exponents and Polish letters have been checked in a rendered sample.
- [ ] No generated graphic contains square fallback glyphs, replacement characters, missing symbols or clipped exponents.
- [ ] Visuals are analytically necessary, not decorative.

### Files and packaging
- [ ] Filenames use the syllabus-derived course prefix.
- [ ] Lecture number is correct.
- [ ] The merged markdown includes the new slides.
- [ ] All referenced image files exist.
- [ ] Zip package contains the merged markdown and all referenced assets.

---

## Failure rules
If any of the following occur, the output is not acceptable and must be corrected before delivery:
- subject-specific hardcoding inside the skill,
- wrong course naming prefix,
- teacher-facing notes on slides,
- teacher-facing teleprompter text,
- unsupported markdown structure,
- prompt text left inside markdown instead of image references,
- missing images,
- non-transparent image backgrounds,
- decorative images with no analytical role,
- overlapping or misaligned diagram elements,
- missing glyphs shown as squares or replacement characters in a generated graphic,
- ambiguous or random examples not supported by syllabus context.

---

## Final instruction to the assistant using this skill
When building a lecture presentation:
1. read the syllabus first,
2. infer the course subject and naming from the syllabus,
3. inspect earlier lecture markdown files for style and continuity,
4. make general, short plan,
5. find papers, literature on topic,
6. download the papers and prepare them for reading and figure extraction,
7. based on literature and sylabus make full presentation plan,
8. make slides in batches of 20,
9. generate PNG assets as separate transparent images,
10. prefer paper-extracted figures when they are suitable, and pad them transparently to fit the slide ratio,
11. reference them using markdown image syntax,
12. merge incrementally after each batch,
13. package the merged markdown and all assets in a zip,
14. run the full quality checklist before delivery.

The result must be reusable across subjects while always adapting precisely to the given syllabus. Typical leccture should be around 120 slides without repetitions. 
