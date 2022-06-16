# python-docx-replace

---

This library was built on top of [python-docx](https://python-docx.readthedocs.io/en/latest/index.html) and the main purpose is to replace words inside a document _**without losing the format**_.

Let's explain the process behind the library:

## First way, losing formatting

One of the ways to replace a key inside a document is by doing something like the code below. Can you do this? YES! But you are going to lose all the paragraph formatting.

```python
key = "${name}"
value = "Ivan"
for p in _get_all_paragraphs(doc):
    if key in p.text:
        p.text = p.text.replace(key, value)
```

## Second way, not all keys

Using the python-docx library, each paragraph has a couple of runs which is a proxy for objects wrapping <w:r> element. We are going to tell more about it later and you can see more details [in the docs](https://python-docx.readthedocs.io/en/latest/api/text.html#run-objects).

You can try replacing the text inside the runs and if it works, then your job is done:

```python
key = "${name}"
value = "Ivan"
for p in _get_all_paragraphs(doc):
    for run in p.runs:
        if key in run.text:
            run.text = run.text.replace(key, value)
```

The problem here is that the key can be broken in more than one run, and then you won't be able to replace it, for example:

It's going to work:

```bash
Word Paragraph: "Hello ${name}, welcome!"
Run1: "Hello ${name}, w"
Run2: "elcome!"
```

It's NOT going to work:

```bash
Word Paragraph: "Hello ${name}, welcome!"
Run1: "Hello ${na"
Run2: "me}, welcome!"
```

You are probably wondering, why does it break paragraph text this way? What are the purpose of the run?

Imagine a Word document with this format:

![word](word.png)

Considering this, what would the format be after parsing the key? Highlighted yellow? Bold and underline? Red with another font? All of them?

That's the purpose of runs, each run hides their sets.

> The final format will be the format that is present in the $ character. All of the others key's characters and their formats will be discarded. In the example above, the final format will be highlighted yellow.

## Solution

The solution adopted is quite simple. First we try to replace in the simplest way, as in the previous example. If it's work, great, all done! If it's not, we build a table of indexes:

```bash
key = "${name}"
value = "Ivan"

Word Paragraph: "Hello ${name}, welcome!"
Run1: "Hello ${na"
Run2: "me}, welcome!"

Word Paragraph: 'H' 'e' 'l' 'l' 'o' ' ' '$' '{' 'n' 'a' 'm' 'e' '}' ',' ' ' 'w' 'e' 'l' 'c' 'o' 'm' 'e' '!'
Char Indexes:    0   1   2   3   4   5   6   7   8   9   10  11  12  13  14  15  16  17  18  19  20  21  22
Run Index:       0   0   0   0   0   0   0   0   0   0   1   1   1   1   1   1   1   1   1   1   1   1   1
Run Char Index:  0   1   2   3   4   5   6   7   8   9   0   1   2   3   4   5   6   7   8   9   10  11  12

Here we have the char indexes, the index of each run by char index and the run char index by run. A little confusing, right? 

With this table we can process and replace all the keys, getting the result:

# REPLACE PROCESS:
Char Index 6 = p.runs[0].text = "Ivan"  # replace '$' by the value
Char Index 7 = p.runs[0].text = ""  # clean all the others parts
Char Index 8 = p.runs[0].text = ""
Char Index 9 = p.runs[0].text = ""
Char Index 10 = p.runs[1].text = ""
Char Index 11 = p.runs[1].text = ""
Char Index 12 = p.runs[1].text = ""
```

After that, we are going to have:

```bash
Word Paragraph: 'H' 'e' 'l' 'l' 'o' ' ' 'Ivan' '' '' '' '' '' '' ',' ' ' 'w' 'e' 'l' 'c' 'o' 'm' 'e' '!'
Indexes:         0   1   2   3   4   5   6      7  8  9 10 11 12  13  14  15  16  17  18  19  20  21  22
Run Index:       0   0   0   0   0   0   0      0  0  0 1  1  1   1   1   1   1   1   1   1   1   1   1
Run Char Index:  0   1   2   3   4   5   6      7  8  9 0  1  2   3   4   5   6   7   8   9   10  11  12
```

All done, now you Word document is fully replaced keeping all the format.

## How to install

### Via PyPI

```bash
pip3 install python-docx-replace
```

### Vanilla

Grab the docx_replace.py file from the src folder and be happy!

## How to use

```python
from python_docx_replace.docx_replace import docx_replace

# get your document using python-docx
doc = Document("document.docx")

# call the replace function with your key value pairs
docx_replace(doc, name="Ivan", phone="+55123456789")

# do whatever you want after that, usually save the document
doc.save("replaced.docx")
```
