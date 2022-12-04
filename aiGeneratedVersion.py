import glob
import urllib.parse
from email import message_from_file

# Get a list of all files with the .msg extension in the current directory
files = glob.glob("*.msg")

# Iterate over the files
for file in files:
    # Open the file and parse it as an email message
    with open(file, "rb") as fp:
        msg = message_from_file(fp)

    # Iterate over the URLs in the message body
    for url in msg.get_payload():
        # Parse the URL using urlparse
        parsed_url = urllib.parse.urlparse(url)

        # Print the URL, without the trailing ">" character
        print(parsed_url.geturl().rstrip(">"))
