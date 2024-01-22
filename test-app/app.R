library(shiny)
library(Microsoft365R)
library(AzureAuth)
library(AzureGraph)
library(stringr)
library(dplyr)
library(purrr)

app      <- "cf81189c-b1be-492e-929e-6e47c3706346"
tenant   <- "ChurchArmy787"
redirect <- "https://church-army.shinyapps.io/FCtest"
resource <- c("https://graph.microsoft.com/.default", "openid")
secret <- readLines("secret/app-secret")


port <- httr::parse_url(redirect)$port
options(shiny.port=if(is.null(port)) 443 else as.numeric(port))

# Define UI for application that draws a histogram
ui <- fluidPage(
  p(""),
  p("Hello", textOutput("display_name", inline = TRUE)),
  p("These are the five largest items in the root of your OneDrive:"),
  verbatimTextOutput("file_list"),
  p(""),
  actionButton("greeting", "Save a friendly greeting to my OneDrive"),
  textOutput("greeting_saved"),
  p(""),
  p("These are the teams that you are part of:"),
  verbatimTextOutput("team_names"),
  verbatimTextOutput("debug")
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

  output$display_name <- renderText("[...]")

  opts <- parseQueryString(isolate(session$clientData$url_search))
  if(is.null(opts$code))
    return()

  ## Authenticate === === === === === === === === === === === === === === ===
  token <- get_azure_token(resource, tenant, app,
                           auth_type = "authorization_code",
                           authorize_args = list(redirect_uri = redirect),
                           version = 2,
                           use_cache = FALSE,
                           auth_code = opts$code,
                           password = secret)

  ## Get user === === === === === === === === === === === === === === === ===
  user <- ms_graph$
    new(token = token)$
    get_user()

  ## Get files === === === === === === === === === === === === === === === ===
  drive <- user$get_drive()
  files <- drive$list_files()

  files <-
    files |>
    arrange(-size) |>
    slice_head(n = 5)

  file_names <- files$name

  output$file_list <- renderPrint(file_names)

  ## Greet user by name === === === === === === === === === === === === === ===
  given_name <- user$properties$givenName
  role       <- user$properties$jobTitle

  output$display_name <- renderText(str_c(given_name, " (", role, ")"))


  ## Print user team names === === === === === === === === === === === === ===
  teams <- user$list_teams()

  output$team_names <- renderPrint({
    map_chr(teams, \(x) pluck(x, "properties", "displayName"))
  })

  # output$debug <- renderPrint(teams_properties)

  observeEvent(input$greeting,
               {
                 tmp <- tempfile(fileext = ".txt")
                 writeLines(str_c("Hello ", given_name,". Have a rewarding and productive day, and don't forget to take some good breaks!"),
                            con = tmp)

                 file_name <- "friendly-greeting-from-shiny-app.txt"
                 drive$upload_file(tmp, dest = file_name)

                 output$greeting_saved <- renderText({
                   str_c("A friendly greeting has been saved!",
                         " It is called '", file_name, "'.")
                   })
               })

  }

# Run the application
shinyApp(ui = ui_func, server = server)
