# How to convert latex to docx

> This was built for latex documents that use the tufte-latex template. All formatting will be removed.

1. Prepare the original _.tex_ document (do a backup first).

Replace all `marginfigure` and `figure*` environments to `figure`; replace
`newthought` to `textsc`; replace `newthought` to `textsc`; and refer the
figures to the explicit relative path where they are.

2. Convert the replaced _.tex_ document to a _.docx_ file.

```bash
pandoc main_replaced.tex -o main.md
pandoc main.md  -o main.docx
```

3. Transform the references to hyperlinks.

```bash
# install the dependencies before running the script
git clone https://github.com/carrascomj/howtolatextodocx.git
cd howtolatextodocx
pip install -r requirements.lock
# run the script
python src/docx_refs/__init__.py PATH/TO/BIBTEX.bib PATH/TO/DOCX.bib PATH/TO/NEW_DOCX.bib
```

4. Tweak final document: replace figures if they were not properly translated; check equations and footnotes (screenshot+paste as last resort).


### FAQ

* **Q: Why not directly translate to `docx` with pandoc?**

Because certain layouts in latex are not reproducible in DOCX and end up as a gigantic mess. 
For sharing and reviewing, it is only possible to do it this way, which will
remove the formatting but will produce something that can be read.
