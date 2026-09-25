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
        S3_DESTINATION = 's3://snapmint-scraper-739589793672-ap-south-1-an/fk_smartphone/'
        "PATH+VENV" = "${WORKSPACE}/.venv/bin"
    }

    stages {
        stage('Prepare Python') {
            steps {
                sh '''
                    set -eux
                    python3 -m venv .venv
                    python -m pip install --upgrade pip
                    pip install -r requirements.txt
                '''
            }
        }

        stage('Install Playwright Browsers') {
            steps {
                sh '''
                    set -eux
                    playwright install --with-deps chromium
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

        stage('Run Scraper') {
            steps {
                sh '''
                    set -eux
                    python smartphone_fk.py
                '''
            }
        }

        stage('Prepare Output') {
            steps {
                sh '''
                    set -eux
                    echo "Files created by scraper:"
                    find . -maxdepth 3 -type f -name "*.xlsx" -print

                    SCRAPE_DATE=$(TZ=Asia/Kolkata date +"%Y-%m-%d")
                    FILE=$(find . -maxdepth 1 -type f -name "flipkart_mobile_*.xlsx" | head -n 1)

                    if [ -z "$FILE" ]; then
                        echo "ERROR: No Flipkart Excel file was created."
                        echo "Files in workspace:"
                        find . -maxdepth 3 -type f
                        exit 1
                    fi

                    echo "Found Excel file: $FILE"
                    mv "$FILE" "flipkart_mobile_assortment_${SCRAPE_DATE}.xlsx"

                    echo "Final file:"
                    ls -lh *.xlsx
                '''
            }
        }

        stage('Verify AWS Identity') {
            steps {
                sh '''
                    set -eux
                    aws sts get-caller-identity
                '''
            }
        }

        stage('Upload Excel to S3') {
            steps {
                sh '''
                    set -eux
                    FILE=$(find . -maxdepth 1 -type f -name "*.xlsx" | head -n 1)

                    if [ -z "$FILE" ]; then
                        echo "ERROR: No Excel file found for S3 upload."
                        exit 1
                    fi

                    echo "Uploading: $FILE"
                    aws s3 cp "$FILE" "$S3_DESTINATION"
                    echo "S3 upload completed successfully."
                '''
            }
        }
    }

    post {
        success {
            archiveArtifacts artifacts: '*.xlsx', allowEmptyArchive: false
        }
    }
}