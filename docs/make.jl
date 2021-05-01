
using Documenter, RateLimiter

makedocs(
    modules = [RateLimiter],
    sitename="RateLimiter.jl", 
    authors = "Chip Kent",
    format = Documenter.HTML(),
)

deploydocs(
    repo = "github.com/chipkent/RateLimiter.jl.git", 
    devbranch = "main",
    push_preview = true,
)
