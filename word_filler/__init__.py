from dataclasses import dataclass
from enum import Enum
from typing import Union, List, Any, Dict, Tuple
from pathlib import Path
import win32com.client

class ContentControlType(Enum):
    RichText = 0
    Text = 1
    Picture = 2
    ComboBox = 3
    DropdownList = 4
    BuildingBlock = 5
    Date = 6
    Group = 7
    Checkbox = 8


@dataclass
class Format:

    tag: str
    start: int
    stop: int
    attrs: dict


def split_text(fmttext: str) -> Tuple[str, List[Format]]:
    text = ""

    curtag = ""
    tagname = False
    tag_stack = []
    text_index = 0
    escaped = False
    for c in fmttext:
        if c == "\\" and not escaped:
            escaped = True
        elif c == "<" and not escaped:
            curtag = ""
            tagname = True
        elif c == ">" and tagname and not escaped:
            tagname, *attrs = curtag.split(" ")
            if tagname.startswith("/"):
                for t in tag_stack[::-1]:
                    if t.tag == tagname[1:]:
                        t.stop = text_index
                        break
                else:
                    raise RuntimeError("Could not find matching start for {curtag} tag")
            else:
                tag_stack.append(
                    Format(tag=tagname, start=text_index, stop=-1, attrs=dict([a.split("=") for a in attrs]))
                )
            tagname = False
        elif tagname:
            curtag += c
            escaped = False
        else:
            text += c
            text_index += 1
            escaped = False
    return text, tag_stack

@dataclass
class ContentControlField:
    """1:n mapping to contentcontrol fields to allow for multiple identical fields.
    """
    type: ContentControlType
    name: str
    objs: List[Any]

    @property
    def value(self):
        values = [o.Range.Text for o in self.objs]
        return values

    @value.setter
    def value(self, text):
        for obj in self.objs:
            if self.type in (ContentControlType.RichText, ContentControlType.Text, ContentControlType.ComboBox):
                valid_text, fmt_controls = split_text(text)
                obj.Range.Text = valid_text
                for fmt in fmt_controls:
                    subrange = obj.Range.Duplicate
                    subrange.SetRange(subrange.start + fmt.start, subrange.start + fmt.stop)
                    if fmt.tag == "b":
                        subrange.Font.Bold = True
                    elif fmt.tag == "i":
                        subrange.Font.Italic = True
                    elif fmt.tag == "font":
                        subrange.Font.Size = int(fmt.attrs['size'])
                    else:
                        raise RuntimeError(f"Unsupported tag: {fmt.tag} for {fmt} when processing {text}")
            elif self.type in (ContentControlType.DropdownList,):
                if self.name == "Patientengeschlecht":
                    if text == "m":
                        obj.DropDownListEntries[1].Select()
                    elif text == "w":
                        obj.DropDownListEntries[0].Select()
                    else:
                        raise RuntimeError(f"Patientengeschlecht only m/w, instead got {text}")
                else:
                    for entry in obj.DropDownListEntries:
                        if str(entry) == text:
                            entry.Select()
                            break
                    else:
                        raise RuntimeError("No matching entry found!")
            elif self.type == ContentControlType.Date:
                obj.Range.Text = text
            else:
                raise RuntimeError(f"{self.name}: Unsupported fieldtype {self.type}")

    def __repr__(self):
        return f"[{self.type.name}]{self.name}|{self.value}|"


class Document:
    def __init__(self, word, doc):
        self._word = word
        self._doc = doc

    @classmethod
    def open(cls, path: Union[str, Path], visible=False) -> "Document":
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = int(visible)
        print(path)
        doc = word.Documents.Open(str(path))
        return cls(word, doc)

    
    def get_fields(self) -> Dict[str, ContentControlField]:
        """Get specified ContentControl fields in the document.
        """
        controls_by_name = {}
        for story in self._doc.StoryRanges:
            while story:
                for control in story.ContentControls:
                    name = control.title
                    if name in controls_by_name:
                        controls_by_name[name].objs.append(control)
                    else:
                        controls_by_name[name] = ContentControlField(
                            type=ContentControlType(control.Type),
                            name = control.title,
                            objs = [control]
                        )
                story = story.NextStoryRange
        return controls_by_name

    def set_fields(self, mapping: Dict[str, str]):
        """Set specified mappings in Document.
        """
        fields = self.get_fields()
        for key, value in mapping.items():
            fields[key].value = value
    
    def save(self, path: Union[str, Path]):
        """Save to new docx document.
        """
        self._doc.SaveAs2(str(path), FileFormat=12)

    def close(self):
        self._doc.Close()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()