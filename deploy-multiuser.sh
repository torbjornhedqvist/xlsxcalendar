#!/bin/bash

# Multi-user XlsxCalendar Deployment Script

set -e

echo "Setting up XlsxCalendar Multi-User Service..."

# Check for generate key flag
if [ "$1" = "--generate-key" ]; then
    echo "Generating new API key..."
    # Generate key directly and add to files
    NEW_KEY=$(openssl rand -base64 32 | tr -d '=' | tr '+/' '-_')
    KEY_HASH=$(echo -n "$NEW_KEY" | sha256sum | cut -d' ' -f1)
    
    # Add to hash file
    echo "$KEY_HASH" >> ~/xlsxcalendar_keys/xlsxcalendar_api_keys
    # Add to plaintext file
    echo "$NEW_KEY" >> ~/xlsxcalendar_keys/xlsxcalendar_api_keys.txt
    
    echo "Generated new API key: $NEW_KEY"
    echo "Key has been saved and will persist across restarts"
    exit 0
fi

# Generate API key if not provided (for first run)
# Note: API keys are now managed internally by the application
echo "API keys are managed automatically by the service"

# Create SSL directory if using SSL
if [ "$1" = "--ssl" ]; then
    mkdir -p ssl
    if [ ! -f ssl/cert.pem ] || [ ! -f ssl/key.pem ]; then
        echo "Generating self-signed SSL certificate..."
        openssl req -x509 -newkey rsa:4096 -keyout ssl/key.pem -out ssl/cert.pem -days 365 -nodes \
            -subj "/C=US/ST=State/L=City/O=Organization/CN=localhost"
    fi
    
    echo "Starting with SSL support..."
    docker-compose -f docker-compose.multiuser.yml --profile ssl up -d
    echo "Service available at: https://localhost"
else
    echo "Starting without SSL..."
    docker-compose -f docker-compose.multiuser.yml up -d
    echo "Service available at: http://localhost:8080"
fi

echo "Deployment complete!"
echo ""
echo "API keys are now persistent across restarts"
echo "To generate additional API keys: ./deploy-multiuser.sh --generate-key"
echo "To stop the service: docker-compose -f docker-compose.multiuser.yml down"
