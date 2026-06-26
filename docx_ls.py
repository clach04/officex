import sys
import zipfile


def get_zip_entries(zf, prefix):
    return [e for e in zf.infolist() if e.filename.startswith(prefix) and not e.filename.endswith('/')]


def print_entries(entries):
    for e in entries:
        y, mo, d = e.date_time[0], e.date_time[1], e.date_time[2]
        print(f"  {e.file_size:>10d}  {y:04d}-{mo:02d}-{d:02d}  {e.filename}")


def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <file.docx>", file=sys.stderr)
        sys.exit(1)

    path = sys.argv[1]
    with zipfile.ZipFile(path) as zf:
        attachments = get_zip_entries(zf, "word/embeddings/")
        if attachments:
            print("Attachments (word/embeddings/):")
            print_entries(attachments)

        media = get_zip_entries(zf, "word/media/")
        if media:
            print("Media (word/media/):")
            print_entries(media)

    if attachments:
        print("WARNING: this document contains embedded attachments.")


if __name__ == "__main__":
    main()
