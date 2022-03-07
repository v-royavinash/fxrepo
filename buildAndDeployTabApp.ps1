# Add command to clone your git repo to local and navigate to repo folder
 git clone https://github.com/v-royavinash/fxrepo.git
 cd \fxrepo

# Following sample script needs to be executed at the root of project folder
# Remember set required environment variables for build first

# Enable static website feature for Azure Storage
#az storage blob service-properties update --account-name '{your_storage_account_name}' --static-website --404-document '{your_404_page_path}' --index-document '{your_index_page_path}' --account-key '{your_storage_account_access_key}'

# Build the tab app
cd tabs
npm install
npm run build

# Remove previous binary
az storage remove --account-name 'mifteocdev9b6e851tab' --container-name '$web' --recursive --account-key 's3sJsnelIxXhGUU4fZQmMpLwXcoAC7y23pppRrnhmdZdR2FGBo08QKQDGyvYx/JdEWe3rrIY+AfQ+AStVoDBtg=='

# Upload new binary
az storage copy --source './build/*' --destination 'https://mifteocdev9b6e851tab.blob.core.windows.net/$web/' --recursive --account-key 's3sJsnelIxXhGUU4fZQmMpLwXcoAC7y23pppRrnhmdZdR2FGBo08QKQDGyvYx/JdEWe3rrIY+AfQ+AStVoDBtg=='