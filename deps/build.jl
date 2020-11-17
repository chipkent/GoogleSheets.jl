
# To run this script:
# pkg> build GoogleSheets

using Conda

Conda.add_channel("conda-forge")
Conda.update()
Conda.add(["google-api-python-client","google-auth-httplib2","google-auth-oauthlib"])
