"""Scan through all .md files in current folder and add support/feedback link to each file.
"""
from pathlib import Path

def main():
    """Make changes to *.md in the current folder.
    """

    # WARNING: we start by deleting *.bak in the current folder.
    for file in Path.cwd().glob('*.bak'):
        file.unlink()

    # rename *.md to *.bak
    #for file in Path.cwd().glob('*.md'):
    for file in Path.cwd().glob('*.md'):
        file.rename(file.with_suffix('.bak'))

    # now scan through *.bak and add the link to each file and write to *.md
    #for file in Path.cwd().glob('*.bak'):
    for file in Path.cwd().glob('*.bak'):
        with open(file.with_suffix('.md'), 'w', encoding='UTF-8') as outfile:
            # Note the use of rstrip() below - we remove any trailing CR/LFs
            # from the source file to assure that there are always exactly two
            # line feeds between the last line of file content and the added link.
            outfile.write( \
                file.read_text(encoding='UTF-8').rstrip() + \
                '\n\n[!include[Support and feedback](~/includes/feedback-boilerplate.md)]')

    # WARNING: we finish by deleting *.bak in the current folder.
    for file in Path.cwd().glob('*.bak'):
        file.unlink()
				
if __name__ == '__main__':
    main()
