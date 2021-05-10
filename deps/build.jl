
# To run this script:
# pkg> build GoogleSheets

using PyCall
@pyimport pip
pip.main(["install","google-api-python-client","google-auth-httplib2","google-auth-oauthlib","FAIL-HARD"])