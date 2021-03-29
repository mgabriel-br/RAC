
### Function to export results from summarize() and assess()
exportToExcel <- function(.object, .file = "results.xlsx") {
  
  ## Install openxlsx if not already installed
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop2(
      "Package `openxlsx` required. Use `install.packages(\"openxlsx\")` and rerun.")
  }
  
  wb <- openxlsx::createWorkbook()
  
  ## Check class
  if(inherits(.object, "cSEMSummarize_default")) {
    est <- .object$Estimates
    ## Add worksheets
    addWorksheet(wb, "Path coefficients")
    addWorksheet(wb, "Loadings")
    addWorksheet(wb, "Weights")
    addWorksheet(wb, "Inner weights")
    addWorksheet(wb, "Residual correlation")
    addWorksheet(wb, "Construct scores")
    addWorksheet(wb, "Indicator correlation matrix")
    addWorksheet(wb, "Composite correlation matrix")
    addWorksheet(wb, "Construct correlation matrix")
    addWorksheet(wb, "Total Effects")
    addWorksheet(wb, "Indirect Effects")
    
    ## Write data to worksheets
    
    writeData(wb, sheet = "Path coefficients", est$Path_estimates)
    writeData(wb, sheet = "Loadings", est$Loading_estimates)
    writeData(wb, sheet = "Weights", est$Weight_estimates)
    writeData(wb, sheet = "Inner weights", est$Inner_weight_estimates)
    writeData(wb, sheet = "Residual correlation", est$Residual_correlation)
    writeData(wb, sheet = "Construct scores", est$Construct_scores)
    writeData(wb, sheet = "Indicator correlation matrix", est$Indicator_VCV)
    writeData(wb, sheet = "Composite correlation matrix", est$Proxy_VCV)
    writeData(wb, sheet = "Construct correlation matrix", est$Construct_VCV)
    writeData(wb, sheet = "Total Effects", est$Effect_estimates$Total_effect)
    writeData(wb, sheet = "Indirect Effects", est$Effect_estimates$Indirect_effect)
    
  } else if(inherits(.object, "cSEMAssess")) {
    
    ## Add worksheets
    addWorksheet(wb, "AVE")
    addWorksheet(wb, "R2")
    addWorksheet(wb, "R2_adj")
    addWorksheet(wb, "Reliability")
    addWorksheet(wb, "Distance and Fit measures")
    addWorksheet(wb, "Model selection criteria")
    addWorksheet(wb, "VIFs")
    addWorksheet(wb, "Effect sizes")
    addWorksheet(wb, "HTMT")
    addWorksheet(wb, "Fornell-Larcker matrix")
    
    ## Write data to worksheets
    writeData(wb, sheet = "AVE", data.frame("Name" = names(.object$AVE), "AVE" = .object$AVE))
    writeData(wb, sheet = "R2", data.frame("Name" = names(.object$R2), "R2" = .object$R2))
    writeData(wb, sheet = "R2_adj", data.frame("Name" = names(.object$R2_adj), "R2_adj" = .object$R2_adj))
    
    # Reliability
    d <- data.frame(
      "Name" = names(.object$Reliability$Cronbachs_alpha),
      "Cronbachs_alpha" = .object$Reliability$Cronbachs_alpha,
      "Joereskogs_rho"  = .object$Reliability$Joereskogs_rho,
      "Dijkstra-Henselers_rho_A" = .object$Reliability$`Dijkstra-Henselers_rho_A`,
      "RhoC" = .object$RhoC,
      "RhoC_mm" = .object$RhoC_mm,
      "RhoC_weighted" = .object$RhoC_weighted,
      "RhoC_weighted_mm" = .object$RhoC_weighted_mm,
      "RhoT" = .object$RhoT,
      "RhoT_weighted" = .object$RhoT_weighted
    )
    writeData(wb, sheet = "Reliability", d)
    
    # Distance and fit
    d <- data.frame(
      "Geodesic distance"          = .object$DG,
      "Squared Euclidian distance" = .object$DL,
      "ML distance"                = .object$DML,
      "Chi_square"                 = .object$Chi_square,
      "Chi_square_df"              = .object$Chi_square_df,
      "CFI"                        = .object$CFI,
      "CN"                         = .object$CN,
      "GFI"                        = .object$GFI,
      "IFI"                        = .object$IFI,
      "NFI"                        = .object$NFI,
      "NNFI"                       = .object$NNFI,
      "RMSEA"                      = .object$RMSEA,
      "RMS_theta"                  = .object$RMS_theta,
      "SRMR"                       = .object$SRMR,
      "Degrees of freedom"         = .object$Df
    )
    writeData(wb, sheet = "Distance and Fit measures", d)
    
    # Model selection criteria
    d <- data.frame(
      "Name" = names(.object$AIC),
      "AIC"  = .object$AIC,
      "AICc" = .object$AICc,
      "AICu" = .object$AICu,
      "BIC"  = .object$BIC,
      "FPE"  = .object$FPE,
      "GM"   = .object$GM,
      "HQ"   = .object$HQ,
      "HQc"  = .object$HQc,
      "Mallows_Cp" = .object$Mallows_Cp
    )
    writeData(wb, sheet = "Model selection criteria", d)
    writeData(wb, sheet = "VIFs", data.frame("Name" = rownames(.object$VIF), .object$VIF))
    writeData(wb, sheet = "Effect sizes", data.frame("Name" = rownames(.object$F2), .object$F2))
    writeData(wb, sheet = "HTMT", .object$HTMT)
    writeData(wb, sheet = "Fornell-Larcker matrix", .object$`Fornell-Larcker`)
    
    
  } else {
    stop2(
      "The following error occured in the exportToExcel() function:\n",
      "`.object` must be a `cSEMSummarize_default` or `cSEMAssess` object."
    )
  }
  
  ## Save workbook
  saveWorkbook(wb, file = .file, overwrite = TRUE)
}

# END --------------------------------------------------------------------------

### Load packages
library(summarytools)  # For descriptive statistics
library(mvnormalTest)  # For normality tests
library(cSEM)          # For SEM 

### Load data
# install.packages("readxl") # Install if not already installed!
database <- readxl::read_xlsx("Surveywikidata20120109.xlsx")

# check missing data
is.na(database) 

# Alternatively (and more conveniently, i think) using vis_dat()
# install.packages("visdat") # Install if not already installed!
visdat::vis_dat(database) # if you had missing values, you'd see them here. 
                          # Check out: https://docs.ropensci.org/visdat/

#summarize data (descriptive statistics)
summarytools::dfSummary(database)

#test univariate and multivariate normality
mvnormalTest::mardia(database, std = "TRUE")
mvnormalTest::msw(database)

# NOTE: it is not surprising that they are not normal! A test does not make much
#       sense since all columns have only a few discrete values. They cannot be normally
#       distributed. 

### defining and assessing the structural equation modeling
model <- '
# define measurement model
AR  =~ AR1 + AR2
COM =~ COM1 + COM2 + COM3 + COM4
PT  =~ PT1 + PT2 + PT3
FC  =~ FC1 + FC2 + FC3
SE  =~ SE1 + SE2 + SE3
PU  =~ PU1 + PU2 + PU3 + PU4
SN  =~ SN1 + SN2
PBC =~ PBC1 + PBC2 + PBC3
BI  =~ BI1 + BI2 + BI3

# define structural model
PU ~ AR + COM + PT
PBC ~ FC + SE
BI ~ PU + SN + PBC
'

#syntax to calculate the structural equation model, as defined previously
result <- csem(.data = database, 
               .model = model, 
               .resample_method = "bootstrap",
               .disattenuate = FALSE, 
               .R = 300, # I reduced to 300 to make it run faster
               .sign_change_option = "individual")

#summary and assessment of structural equation model results
summarized <- summarize(result)
assessed <- assess(result)

# Export results from summarize()
exportToExcel(summarized, .file = "results_summarize.xlsx")
# Export results from assess()
exportToExcel(assessed, .file = "results_assess.xlsx")


