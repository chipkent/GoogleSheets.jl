
using Documenter, GoogleSheets

makedocs(
    modules = [GoogleSheets],
    sitename="GoogleSheets.jl", 
    authors = "Chip Kent",
    format = Documenter.HTML(),
)

deploydocs(
    repo = "github.com/chipkent/GoogleSheets.jl.git", 
    devbranch = "main",
    push_preview = true,
)
