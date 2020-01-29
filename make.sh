if [ ! -e ./custom-reference.docx ]; then
    # https://pandoc.org/MANUAL.html#options-affecting-specific-writers
    pandoc -o custom-reference.docx --print-default-data-file reference.docx
fi

pandoc --reference-doc custom-reference.docx -o zoyo.docx zoyo.md
