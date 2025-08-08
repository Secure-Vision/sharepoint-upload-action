# Dockerfile
# Use a slim, official Python image
FROM python:3.9-slim

# Copy all the files from your repository into the container's root
COPY . /

# Install the required Python libraries
RUN pip install requests msal pathspec

# Make the entrypoint script executable
RUN chmod +x /entrypoint.sh

# Set the entrypoint for the container
ENTRYPOINT ["/entrypoint.sh"]