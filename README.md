# clip2png #

**clip2png** is a [Microsoft JScript](https://en.wikipedia.org/wiki/JScript) that is designed to help simplify the job of collecting and organizing notes about how to accomplish a particular task on a Windows Platform.

**clip2png** detects whether you have an picture on your Microsoft Windows clipboard.  If there is an picture on the clipboard, it dumps the clipboard to a PNG file, and helps you capture some basic metadata about why the image is interesting.

The script outputs:

1. A collection of images, organized in chronological order.
2. A simple JSON file

## Requirements ##

**clip2png** requires Microsoft Word to be installed on the system, in order to run.  It should run on any modern Microsoft Windows platform, going as far back as Windows Vista.

## Configuration ##

In order to run **clip2png**, you should first use a text editor to modify **settings.json**:

```json
{
"fullPath": "C:|Users|Chris|Desktop",
"projectName": "test_project"
}
```

You should change the value of "fullPath" to point to a folder where you want **clip2png** to dump its output.

The value for "fullPath" should use vertical bars instead of the traditional backslashes that are used in Microsoft PATHs.  (In JSON, backslashes must be escaped with another backslash.  For convenience and brevity, **clip2png** automatically maps vertical bars to the proper backslash character.)

**clip2png** will create a new folder in "fullPath" named whatever the value of "projectName" is.

## Using clip2png ##

Simply double click on **clip2png.js**.
