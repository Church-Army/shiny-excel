library(shiny)
library(Microsoft365R)
library(AzureAuth)
library(AzureGraph)

app      <- "cf81189c-b1be-492e-929e-6e47c3706346"
tenant   <- "ChurchArmy787"
redirect <- "https://church-army.shinyapps.io/FCtest"
resource <- c("https://graph.microsoft.com/.default", "openid")

port <- httr::parse_url(redirect)$port
options(shiny.port=if(is.null(port)) 443 else as.numeric(port))

# Define UI for application that draws a histogram
ui <- fluidPage(
  textOutput("file_list")
)

ui_func <- function(req)
{
  opts <- parseQueryString(req$QUERY_STRING)
  if(is.null(opts$code))
  {
    auth_uri <- build_authorization_uri(resource, tenant, app,
                                        redirect_uri = redirect, version = 2)
    redir_js <- sprintf("location.replace(\"%s\");", auth_uri)
    tags$script(HTML(redir_js))
  }
  else ui
}

# Define server logic required to draw a histogram
server <- function(input, output, session) {


  opts <- parseQueryString(isolate(session$clientData$url_search))
  if(is.null(opts$code))
    return()

  token <- get_azure_token(resource, tenant, app,
                           auth_type = "authorization_code",
                           authorize_args = list(redirect_uri = redirect),
                           version = 2,
                           use_cache = FALSE,
                           auth_code = opts$code)

  drv <- ms_graph$
    new(token = token)$
    get_user()$
    get_drive()

  files <- drv$list_files()

  file_names <- files$name


  output$file_list <- renderText(file_names[1:10])

  }

# Run the application
shinyApp(ui = ui_func, server = server)
