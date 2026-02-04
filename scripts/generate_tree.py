import os
import argparse

def generate_tree(dir_path, prefix="", exclude_dirs=None):
    if exclude_dirs is None:
        exclude_dirs = []

    tree_str = ""
    
    try:
        entries = os.listdir(dir_path)
    except PermissionError:
        return ""

    entries = sorted([e for e in entries if e not in exclude_dirs])
    
    for i, entry in enumerate(entries):
        path = os.path.join(dir_path, entry)
        is_last = i == len(entries) - 1
        
        connector = "└── " if is_last else "├── "
        
        display_name = entry
        if os.path.isdir(path):
            display_name += "/"
        tree_str += f"{prefix}{connector}{display_name}\n"
        
        if os.path.isdir(path):
            extension = "    " if is_last else "│   "
            tree_str += generate_tree(path, prefix + extension, exclude_dirs)
            
    return tree_str

def generate_paths_list(startpath, exclude_dirs=None):
    if exclude_dirs is None:
        exclude_dirs = []
    
    paths_list = []
    
    for root, dirs, files in os.walk(startpath):
        dirs[:] = [d for d in dirs if d not in exclude_dirs]
        dirs.sort()
        
        if root != startpath:
            paths_list.append(os.path.abspath(root) + os.sep)

        for f in files:
            paths_list.append(os.path.abspath(os.path.join(root, f)))
            
    paths_list.sort()
    return "# Project Paths\n\n```txt\n" + "\n".join(paths_list) + "\n```\n"

def generate_file_contents(startpath, exclude_dirs=None):
    if exclude_dirs is None:
        exclude_dirs = []
    
    content_str = "# File Contents\n\n"
    
    for root, dirs, files in os.walk(startpath):
        dirs[:] = [d for d in dirs if d not in exclude_dirs]
        dirs.sort()
        files.sort()

        for f in files:
            if f in ["package-lock.json", "yarn.lock", "pnpm-lock.yaml", "bun.lockb"]:
                continue

            file_path = os.path.join(root, f)
            abs_path = os.path.abspath(file_path)
            
            content_str += f"## File: {abs_path}\n\n"
            
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    file_content = file.read()
                
                _, ext = os.path.splitext(f)
                lang = ext.lstrip(".") if ext else "txt"
                if not lang: lang = "txt"
                
                content_str += f"```{lang}\n{file_content}\n```\n\n"
            except (UnicodeDecodeError, PermissionError):
                content_str += "> [Binary or Non-UTF-8 file content not shown]\n\n"
            except Exception as e:
                content_str += f"> [Error reading file: {str(e)}]\n\n"
                
    return content_str

def main():
    parser = argparse.ArgumentParser(description="Generate project structure.")
    parser.add_argument("-p", action="store_true", help="Generate list of absolute paths only.")
    parser.add_argument("-tp", action="store_true", help="Generate both tree structure and list of absolute paths.")
    parser.add_argument("-c", action="store_true", help="Generate combined output in a single file.")
    parser.add_argument("--content", action="store_true", help="Include file contents in the output.")
    parser.add_argument("--print", action="store_true", help="Print output to terminal.")
    args = parser.parse_args()

    root_dir = os.getcwd()
    docs_dir = os.path.join(root_dir, "docs")
    
    tree_output_file = os.path.join(docs_dir, "project_structure.md")
    paths_output_file = os.path.join(docs_dir, "project_paths.md")
    combined_output_file = os.path.join(docs_dir, "project_combined.md")
    contents_output_file = os.path.join(docs_dir, "project_contents.md")
    
    exclude_dirs = ["node_modules", ".git", "build", "gist", "__pycache__", "dist", ".DS_Store", "Thumbs.db"]

    if not os.path.exists(docs_dir):
        os.makedirs(docs_dir)
        
    generate_tree_structure = True
    generate_paths = False

    if args.tp:
        generate_tree_structure = True
        generate_paths = True
    elif args.p:
        generate_tree_structure = False
        generate_paths = True
    elif args.c:
        generate_tree_structure = False
        generate_paths = False

    if args.c:
        tree = "# Project Structure\n\n```txt\n.\n" + generate_tree(root_dir, "", exclude_dirs) + "```\n"
        paths = generate_paths_list(root_dir, exclude_dirs)
        combined = tree + "\n" + paths
        if args.content:
            combined += "\n" + generate_file_contents(root_dir, exclude_dirs)
        
        with open(combined_output_file, "w", encoding="utf-8") as f:
            f.write(combined)
        if args.print:
            print("\n" + "="*40 + "\nCOMBINED OUTPUT PREVIEW\n" + "="*40 + "\n" + combined)
            
    elif args.content:
        content = generate_file_contents(root_dir, exclude_dirs)
        with open(contents_output_file, "w", encoding="utf-8") as f:
            f.write(content)
        if args.print:
            print("\n" + "="*40 + "\nCONTENTS PREVIEW\n" + "="*40 + "\n" + content)
            
    if generate_tree_structure:
        tree = "# Project Structure\n\n```txt\n.\n" + generate_tree(root_dir, "", exclude_dirs) + "```\n"
        with open(tree_output_file, "w", encoding="utf-8") as f:
            f.write(tree)
        if args.print:
            print("\n" + "="*40 + "\nTREE STRUCTURE PREVIEW\n" + "="*40 + "\n" + tree)
            
    if generate_paths:
        paths = generate_paths_list(root_dir, exclude_dirs)
        with open(paths_output_file, "w", encoding="utf-8") as f:
            f.write(paths)
        if args.print:
            print("\n" + "="*40 + "\nPATHS LIST PREVIEW\n" + "="*40 + "\n" + paths)

if __name__ == "__main__":
    main()
