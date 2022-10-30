# python-docx-replace

---

This library was built on top of [python-docx](https://python-docx.readthedocs.io/en/latest/index.html) and the main purpose is to replace words inside a document _**without losing the format**_.

There is also a functionality that allows defining blocks in the Word document and set if they will be removed or not.

## Replacing a word - docx_replace

You can define a key in your Word document and set the value to be replaced. This program requires the following key format: `${key_name}`

Let's explain the process behind the library:

### First way, losing formatting

One of the ways to replace a key inside a document is by doing something like the code below. Can you do this? YES! But you are going to lose all the paragraph formatting.

```python
key = "${name}"
value = "Ivan"
for p in get_all_paragraphs(doc):
    if key in p.text:
        p.text = p.text.replace(key, value)
```

### Second way, not all keys

Using the python-docx library, each paragraph has a couple of `runs` which is a proxy for objects wrapping `<w:r>` element. We are going to tell more about it later and you can see more details [in the python-docx docs](https://python-docx.readthedocs.io/en/latest/api/text.html#run-objects).

You can try replacing the text inside the runs and if it works, then your job is done:

```python
key = "${name}"
value = "Ivan"
for p in get_all_paragraphs(doc):
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

You are probably wondering, why does it break paragraph text this way? What are the purpose of the `run`?

Imagine a Word document with this format:

![word](word.png)

Each `run` holds their own format! That's the goal for the `runs`.

Considering this and using this library, what would be the format after parsing the key? Highlighted yellow? Bold and underline? Red with another font? All of them?

> The final format will be the format that is present **in the $ character**. All of the others key's characters and their formats will be discarded. In the example above, the final format will be **highlighted yellow**.

### Solution

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

## Replace blocks - docx_blocks

You can define a block in your Word document and set if it is going to be removed or not. The format required for key blocks are exactly like tags `HTML`, as following:

- Initial of block: `<signature>`
- End of the block: `</signature>`

Let's say you define two blocks like this:

Word document:
```bash
Contract

Detais of the contract

<signature>
Please, put your signature here: _________________
</signature>
```

### Setting signature to be removed

```python
docx_blocks(doc, signature=True)
```

Final Word document:
```bash
Contract

Detais of the contract


Please, put your signature here: _________________
```

### Setting signature to not be removed

```python
docx_blocks(doc, signature=False)
```

Final Word document:
```bash
Contract

Detais of the contract

```

### docx_blocks limitation

If there are **tables** inside a block that is set to be removed, these tables are not going to be removed. Tables are different objects in python-docx library and they are not present in the paragraph object.

You can use the function `docx_remove_table` to remove tables from the Word document by their index.

```python
docx_remove_table(doc, 0)
```

> The table index works exactly like any indexing property. It means if you remove an index, it will affect the other indexes. For example, if you want to remove the first two tables, you can't do like this:

```python
docx_remove_table(doc, 0)
docx_remove_table(doc, 1)  # it will raise an index error
```

> You should instead do like this:

```python
docx_remove_table(doc, 0)
docx_remove_table(doc, 0)
```

## How to install

### Via PyPI

```bash
pip3 install python-docx-replace
```

## How to use

```python
from python_docx_replace import docx_replace

# get your document using python-docx
doc = Document("document.docx")

# call the replace function with your key value pairs
docx_replace(doc, name="Ivan", phone="+55123456789")

# call the blocks function with your sets
docx_blocks(doc, signature=True, table_of_contents=False)

# remove the first table in the Word document
docx_remove_table(doc, 0)

# do whatever you want after that, usually save the document
doc.save("replaced.docx")
```

> TIP: If you want to call with a defined `dict` variable, you can leverage the `**` syntax from python:

```python
my_dict = {
    "name": "Ivan",
    "phone": "+55123456789"
}

docx_replace(doc, **my_dict)
```
