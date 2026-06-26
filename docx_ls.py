import sys
import zipfile


def list_zip_entries(zf, prefix):
    entries = [e for e in zf.infolist() if e.filename.startswith(prefix) and not e.filename.endswith('/')]
    for e in entries:
        y, mo, d = e.date_time[0], e.date_time[1], e.date_time[2]
        date_str = f"{y:04d}-{mo:02d}-{d:02d}"
        print(f"  {e.file_size:>10d}  {date_str}  {e.filename}")
    return entries


def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <file.docx>", file=sys.stderr)
        sys.exit(1)

    path = sys.argv[1]
    with zipfile.ZipFile(path) as zf:
        print("Attachments (word/embeddings/):")
        attachments = list_zip_entries(zf, "word/embeddings/")

        print("Media (word/media/):")
        list_zip_entries(zf, "word/media/")

    if attachments:
        print("WARNING: this document contains embedded attachments.")


if __name__ == "__main__":
    main()
