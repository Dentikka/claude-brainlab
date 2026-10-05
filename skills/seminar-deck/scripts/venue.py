"""Venue leads for arXiv papers in one request: the arXiv 'comments' and 'journal_ref' fields
often say "Accepted to ICLR 2026". A comment is a lead, not proof: confirm the venue on the
conference site (proceedings, iclr.cc/virtual, neurips.cc, aclanthology.org, mlanthology.org).
No confirmed venue -> the citation carries only the year.

    python venue.py 2509.20317 2506.18582 2505.18454
"""
import sys
import urllib.request
import xml.etree.ElementTree as ET

NS = {"a": "http://www.w3.org/2005/Atom", "x": "http://arxiv.org/schemas/atom"}


def main(ids):
    url = "http://export.arxiv.org/api/query?max_results=100&id_list=" + ",".join(ids)
    root = ET.fromstring(urllib.request.urlopen(url, timeout=60).read())
    for e in root.findall("a:entry", NS):
        aid = e.find("a:id", NS).text.rsplit("/", 1)[-1]
        title = " ".join(e.find("a:title", NS).text.split())
        year = e.find("a:published", NS).text[:4]
        c, j = e.find("x:comment", NS), e.find("x:journal_ref", NS)
        print(f"{aid} | {year} | {title[:80]}\n    comment: {c.text.strip() if c is not None else '-'}"
              f"\n    journal_ref: {j.text.strip() if j is not None else '-'}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1:])
