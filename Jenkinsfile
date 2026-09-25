pipeline {
    agent any

    triggers {
        cron('TZ=Asia/Kolkata\n0 8 * * *')
    }

    options {
        timestamps()
        buildDiscarder(logRotator(numToKeepStr: '30', artifactNumToKeepStr: '30'))
    }

    environment {
        AWS_DEFAULT_REGION = 'ap-south-1'
        S3_BUCKET = 'snapmint-s3-to-warehouse'
        FLIPKART_DATASET = 'flipkart_mobile_scraper'
        AMAZON_DATASET = 'amazon_mobile_scraper'
        VENV_DIR = '.venv'
    }

    stages {
        stage('Prepare Python') {
            steps {
                sh '''
                    set -eux
                    python3 -m venv "$VENV_DIR"
                    "$VENV_DIR/bin/python" -m pip install --upgrade pip
                    "$VENV_DIR/bin/pip" install -r requirements.txt
                '''
            }
        }

        stage('Install Playwright Browsers') {
            steps {
                sh '''
                    set -eux
                    "$VENV_DIR/bin/playwright" install --with-deps chromium
                '''
            }
        }

        stage('Test Flipkart Connectivity') {
            steps {
                sh '''
                    set -eux
                    curl -I -L --max-time 20 https://www.flipkart.com
                '''
            }
        }

        stage('Run Flipkart Scraper') {
            steps {
                sh '''
                    set -eux
                    "$VENV_DIR/bin/python" smartphone_fk.py
                '''
            }
        }

        stage('Run Amazon Scraper') {
            steps {
                sh '''
                    set -eux
                    "$VENV_DIR/bin/python" smartphone_az.py
                '''
            }
        }

        stage('Prepare Output') {
            steps {
                sh '''
                    set -eux
                    echo "Files created by scraper:"
                    find . -maxdepth 3 -type f -name "*.xlsx" -print

                    SCRAPE_TIMESTAMP=$(TZ=Asia/Kolkata date +"%Y-%m-%d_%H-%M-%S")
                    OUTPUT_DIR="output"
                    rm -rf "$OUTPUT_DIR"
                    mkdir -p "$OUTPUT_DIR/$FLIPKART_DATASET" "$OUTPUT_DIR/$AMAZON_DATASET"

                    convert_excel_to_csv() {
                        dataset="$1"
                        pattern="$2"
                        label="$3"
                        file=$(find . -maxdepth 1 -type f -name "$pattern" | head -n 1)

                        if [ -z "$file" ]; then
                            echo "ERROR: No $label Excel file was created."
                            echo "Files in workspace:"
                            find . -maxdepth 3 -type f
                            exit 1
                        fi

                        target="$OUTPUT_DIR/$dataset/${SCRAPE_TIMESTAMP}.csv"
                        echo "Converting $label Excel file: $file to $target"
                        "$VENV_DIR/bin/python" -c 'import csv, sys; from openpyxl import load_workbook; source_path, target_path = sys.argv[1], sys.argv[2]; workbook = load_workbook(source_path, data_only=True, read_only=True); worksheet = workbook.active; csv_file = open(target_path, "w", newline="", encoding="utf-8"); writer = csv.writer(csv_file); [writer.writerow(["" if value is None else value for value in row]) for row in worksheet.iter_rows(values_only=True)]; csv_file.close(); workbook.close()' "$file" "$target"
                    }

                    convert_excel_to_csv "$FLIPKART_DATASET" "flipkart_mobile_*.xlsx" "Flipkart"
                    convert_excel_to_csv "$AMAZON_DATASET" "amazon_mobile_*.xlsx" "Amazon"

                    echo "Final files:"
                    ls -lh *.xlsx "$OUTPUT_DIR"/*/*.csv
                '''
            }
        }

        // stage('Verify AWS Identity') {
        //     steps {
        //         sh '''
        //             set -eux
        //             aws sts get-caller-identity
        //         '''
        //     }
        // }

        stage('Upload CSV to S3') {
            steps {
                sh '''
                    set -eux
                    FILES=$(find output -mindepth 2 -maxdepth 2 -type f -name "*.csv")

                    if [ -z "$FILES" ]; then
                        echo "ERROR: No CSV file found for S3 upload."
                        exit 1
                    fi

                    for FILE in $FILES; do
                        DATASET=$(basename "$(dirname "$FILE")")
                        BASENAME=$(basename "$FILE")
                        TIMESTAMP=${BASENAME%.csv}
                        YEAR=$(printf "%s" "$TIMESTAMP" | cut -d- -f1)
                        MONTH=$(printf "%s" "$TIMESTAMP" | cut -d- -f2)
                        DAY=$(printf "%s" "$TIMESTAMP" | cut -d- -f3 | cut -d_ -f1)
                        S3_DESTINATION="s3://${S3_BUCKET}/${DATASET}/${YEAR}/${MONTH}/${DAY}/${BASENAME}"

                        echo "Uploading: $FILE to $S3_DESTINATION"
                        aws s3 cp "$FILE" "$S3_DESTINATION"
                    done
                    echo "S3 upload completed successfully."
                '''
            }
        }
    }

    post {
        success {
            archiveArtifacts artifacts: '*.xlsx,output/**/*.csv', allowEmptyArchive: false
        }
    }
}
