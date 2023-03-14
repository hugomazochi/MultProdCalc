library(shiny)
library(openxlsx)
library(shinydashboard)


ui = dashboardPage(
  dashboardHeader(
    title = "CÁLCULO DE MÚLTIPLOS PRODUTOS",
    titleWidth = 450
  ),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Tabela", tabName = "tabela", icon = icon("table-list")),
      menuItem("Gráfico", tabName = "grafico", icon = icon("chart-column")),
      menuItem("Sobre", tabName = "sobre", icon = icon("info"))
    ),
    width = 350,
    br(),
    fileInput('file1', label = h4(strong( 'BASE DE DADOS DO INVENTÁRIO:')),
              accept = c(".xlsx")),
    h4("  Insira um arquivo .xlsx que contenha a base", align="center", color = "white"),
    h4(" de dados do inventario com as colunas ", align="center",  color = "white"),
    h4("referentes as árvores, ao DAP e a HT", align="center",  color = "white"),
    textInput("tab1", label = h4(strong("Nome da Planilha (Inventário):")), "Planilha1"),
    br(),
    br(),
    fileInput('file2', label = h4(strong('BASE DE DADOS DOS PRODUTOS:')),
              accept = c(".xlsx")),
    h4("  Insira um arquivo .xlsx que contenha a base", align="center",  color = "white"),
    h4("de dados dos produtos e seus respectivos", align="center",  color = "white"),
    h4("diâmetros e comprimentos", align="center",  color = "white"),
    textInput("tab2", label = h4(strong("Nome da Planilha (Produtos):")), "Planilha1"),
    br(),
    br(),
    numericInput("b0", "Valor de b0:", value = 1.17885),
    numericInput("b1", "Valor de b1:", value = -4.20444),
    numericInput("b2", "Valor de b2:", value = 19.49670),
    numericInput("b3", "Valor de b3:", value = -42.57810),
    numericInput("b4", "Valor de b4:", value = 40.48840),
    numericInput("b5", "Valor de b5:", value = -14.39020),
    numericInput("htoco", "Altura do toco:", value = 0.1),
    mainPanel(
    )
    
    
  ),
  dashboardBody(
    tags$head(tags$style(HTML('
        .skin-blue .main-header .logo {
          background-color: #3c8dbc;
        }
        .skin-blue .main-header .logo:hover {
          background-color: #3c8dbc;
        }
      '))),
    img(src = "logo_da_ufla.png", align = "center", height = 105, width = 270),
    img(src = "blank.png", height = 20, width = 20),
    img(src = "lemaf.png", height = 105, width = 105),
    br(),
    br(),
    br(),
    tabItems(
      tabItem(tabName = "tabela",
              mainPanel(
                style = "max-height:700px;overflow-y:auto;",
                tableOutput("contents"),
                downloadButton("downloadData", "Baixar Tabela")
              )
      ),
      tabItem(tabName = "grafico",
              mainPanel(
                plotOutput("plot")
              )),
      tabItem(tabName = "sobre",
              mainPanel(
                style = "max-height:700px;overflow-y:auto;",
                h1(strong("Sobre:")),
                h3("Este aplicativo foi desenvolvido para auxiliar no cálculo de múltiplos produtos da madeira. A função de afilamento utilizada para o cálculo dos diâmetros em diferentes alturas da árvore foi o Polinômio de Quinto Grau:"),
                br(),
                img(src = "polinomio.png", align = "center", height = 262, width = 525),
                h3(strong("Observações do funcionamento:")),
                h3("A base de dados inseridas no programa devem conter cabeçalho, o DAP deve estar em centímetros e a HT em metros, o mesmo é válido para os produtos (diâmetro em centímetros e comprimento em metros)."),
                h3("A primeira coluna da tabela do inventário deve ser composta apenas por valores numéricos."),
                br(),
                h3("As imagens a seguir exemplificam a estrutura das bases de dados a serem inseridas:"),
                img(src = "base_inventario.png", align = "center", height = 261, width = 261),
                img(src = "base_produtos.png", align = "center", height = 150, width = 261),
                br(),
                br(),
                h4(" Essa ferramenta foi apresentada por Hugo Mazochi Barroso como parte do seu Trabalho de Conclusão de Curso em Engenharia Florestal na Universidade Federal de Lavras. Os Professores Dr. Samuel José Silva Soares da Rocha e Dr. Lucas Rezende Gomide prestaram orientação e ajuda na elaboração do trabalho e escrita do código."),
                h4("Não nos responsabilizamos pelo mal uso da ferramenta, não oferecemos garantias de qualquer tipo em relação à adequação, confiabilidade, disponibilidade ou precisão do código, e não nos responsabilizamos por quaisquer perdas e danos associados ao seu uso.")
              )))))



server = function(input, output,session){
  output$contents <- renderTable({
    inFile <- input$file1
    if(is.null(inFile))
      return(NULL)
    file.rename(inFile$datapath,paste(inFile$datapath, ".xlsx", sep=""))
    marvore<-read.xlsx(paste(inFile$datapath, ".xlsx", sep=""), sheet=input$tab1)
    
    inFile2 <- input$file2
    if(is.null(inFile2))
      return(NULL)
    file.rename(inFile2$datapath,paste(inFile2$datapath, ".xlsx", sep=""))
    mprod<-read.xlsx(paste(inFile2$datapath, ".xlsx", sep=""), sheet=input$tab2)
    
    marv<-as.matrix(marvore)
    
    b0<-input$b0
    b1<-input$b1
    b2<-input$b2
    b3<-input$b3
    b4<-input$b4
    b5<-input$b5
    htoco<- input$htoco
    
    msaida<-matrix(nrow=nrow(marvore), ncol= 2+nrow(mprod))
    
    #Cálculo do afilamento ---------------------------------------------------
    
    #Passando pelo arquivo de árvores
    for(i in 1:nrow(marv)){
      msaida[i,1]<- i #Salvando o valor da árvore
      hi<-htoco
      #Passando pelos produtos
      for(j in 1:nrow(mprod)){
        ctora<-0
        repeat{
          #Altura na árvore a investigar o di
          hi<-hi+mprod[j,3] 
          #di estimado
          di<-marv[i,2]*(b0+(b1*(hi/marv[i,3])^1)+(b2*(hi/marv[i,3])^2)+(b3*(hi/marv[i,3])^3)+(b4*(hi/marv[i,3])^4)+(b5*(hi/marv[i,3])^5))
          if(di>=mprod[j,2]){
            ctora<-ctora+1  
          }
          else{
            #Voltando na posição anterior de hi
            hi<-hi-mprod[j,3]
            #Salvando as toras do produto j
            msaida[i,1+j]<- ctora
            break
          }
        }
      }
      #Resgate do Valor do Resíduo
      msaida[i, ncol(msaida)]<-marv[i,3]-hi
      msaida<-as.data.frame(msaida)
      num_cols <- ncol(msaida)
      colnames(msaida) <- c("Árvore", paste( mprod[seq(1, num_cols-2),1]), "Resíduo")
    }
    
    output$plot <- renderPlot({
      num_cols <- ncol(msaida)
      col_sums <- sapply(2:(num_cols-1), function(i) sum(msaida[,i]))
      barplot(height = col_sums, 
              names.arg = paste("Produto", mprod[seq(1, num_cols-2),1]), 
              main = "Multiprodutos", 
              xlab = "Produtos", 
              ylab = "Número de Toras",
              col = "gray")
    })
    output$downloadData <- downloadHandler(
      filename = function() {
        paste("tableOutput-", Sys.Date(), ".csv", sep="")
      },
      content = function(file) {
        write.csv(msaida, file, row.names=FALSE)
      }
    )
    
    substituir_coluna <- function(msaida, marvore) {
      msaida[,1] <- marvore[,1]
    }
    msaida_nova <- as.data.frame(msaida)
    msaida_nova[,1] <- as.factor(marvore[,1])
    return(msaida_nova)
    
  })
  
}

shinyApp(ui = ui, server = server)
