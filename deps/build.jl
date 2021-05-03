
# To run this script:
# pkg> build GoogleSheets

using Conda

Conda.add_channel("conda-forge")
Conda.update()
Conda.add(["google-api-python-client","google-auth-httplib2","google-auth-oauthlib"])

println("ENV: ", env)
println(Conda.list())
using PyCall
pickle = pyimport("pickle")
os_path = pyimport("os.path")
build = pyimport("googleapiclient.discovery").build
InstalledAppFlow = pyimport("google_auth_oauthlib.flow").InstalledAppFlow
Request = pyimport("google.auth.transport.requests").Request
open = pybuiltin("open")
println("EVERYTHING WORKED")
