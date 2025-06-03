# AWS Bedrock PPT Translator

Translate PowerPoint (.pptx) slides files using AWS Bedrock LLM models.

## Installation

```bash
$ pip install -r requirements.txt
# $ uv pip install -r requirements.txt
```

## Key Features

- **GenAI translation** using AWS Bedrock LLMs
- **Parallel processing** of slides for faster translations
- **Comprehensive coverage** - translates text in text frames, tables, and slide notes
- **Model flexibility** - support for different AWS Bedrock models
- **Automatic output naming** - creates a new file with target language code appended

## Usage

```bash
python pptx-translator-bedrock.py source_language_code target_language_code input_file_path
```

### Examples

Translate a presentation from English to Spanish:
```bash
python pptx-translator-bedrock.py en es presentation.pptx
```

Use a specific AWS Bedrock model:
```bash
python pptx-translator-bedrock.py en fr presentation.pptx --model us.amazon.nova-pro-v1:0
```

Set AWS region and increase parallel workers:
```bash
python pptx-translator-bedrock.py en de presentation.pptx --region us-west-2 --workers 8
```

For more information on available options:
```bash
python pptx-translator-bedrock.py --help
```

## Command-line Arguments

```
usage: PowerPoint Translator using AWS Bedrock
       [-h] [--model MODEL] [--region REGION] [--workers WORKERS]
       source_lang target_lang file

positional arguments:
  source_lang        Source language code (e.g., en)
  target_lang        Target language code (e.g., es)
  file               PowerPoint file to translate

optional arguments:
  -h, --help         show this help message and exit
  --model MODEL      Bedrock model ID (default: us.amazon.nova-lite-v1:0)
  --region REGION    AWS region name
  --workers WORKERS  Number of parallel workers (default: 4)
```

## How It Works

The translator processes PowerPoint files by:

1. Loading the presentation using the python-pptx library
2. Creating a connection to AWS Bedrock runtime
3. Processing slides in parallel using a thread pool for improved performance
4. Translating text within:
   - Text frames (shapes containing text)
   - Tables (all cell contents)
   - Slide notes
5. Saving the translated presentation with the target language code appended to the filename

The translation leverages AWS Bedrock LLM models, which offer higher quality translations compared to traditional machine translation services.
