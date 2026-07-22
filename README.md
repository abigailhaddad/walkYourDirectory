```markdown
# walkYourDirectory

walkYourDirectory is a powerful tool designed to extract text and metadata from all files within a specified directory, including all its subdirectories. This utility is perfect for data analysis, organization, and management tasks that require comprehensive file information retrieval.

## Features

- **Recursive Directory Traversal**: Automatically navigate through folders and subfolders to access all files.
- **Text Extraction**: Retrieve text content from a wide variety of file formats.
- **Metadata Extraction**: Gather metadata information such as file size, creation date, and modification date.
- **Customizable**: Easily configure the types of files and metadata you want to extract.

## Installation

To use walkYourDirectory, clone the repository and install the necessary dependencies:

```bash
git clone https://github.com/yourusername/walkYourDirectory.git
cd walkYourDirectory
pip install -r requirements.txt
```

## Usage

Here's a simple example of how to use walkYourDirectory:

```python
from walkYourDirectory import DirectoryWalker

# Initialize with the path to the directory you want to scan
walker = DirectoryWalker('/path/to/your/folder')

# Extract text and metadata
files_info = walker.extract_info()

# Print extracted information
for file in files_info:
    print(file)
```

## Contribution

Contributions are welcome! Please follow these steps:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/YourFeature`).
3. Commit your changes (`git commit -am 'Add new feature'`).
4. Push to the branch (`git push origin feature/YourFeature`).
5. Open a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.
```