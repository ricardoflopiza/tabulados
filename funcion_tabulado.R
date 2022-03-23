library(tidyverse)
library(haven)
library(survey)
library(calidad)
library(openxlsx)




#enusc_prueba = readRDS("/home/ricardo/Documents/INE/servicios_compartidos/shiny_calidad/bkish_2019_mas_variables.rds")
enusc2020 = read_sav("base_usuario_ENUSC2020.sav")

#enusc2020$G2_1_1 = dplyr::if_else(enusc2020$G2_1_1==1,1,0) 
#enusc2020$G1_1_1 = dplyr::if_else(enusc2020$G2_1_1==1,1,0) 
#
### filtramos por kish ####
#enusc2020 = enusc2020[enusc2020$Kish == 1,]
#
#dcHOG = svydesign(ids =~Conglomerado, strata = ~VarStrat, weights = ~Fact_Hog ,data = enusc2020)
#options(survey.lonely.psu = "certainty")
#
#create_prop(G2_1_1, disenio = dcHOG, subpop = G1_1_1)

plan2020 = read.xlsx("plan_tabulado2020.xlsx")
plan2019 = read.xlsx("plan_tabulado.xlsx")
llave = read.xlsx("/home/ricardo/Documents/INE/ENUSC/Enusc2020/bases/nombres_estandar_enusc.xlsx")  

plan2020$Cuadro = stringr::str_remove(tolower(plan2020$Cuadro),".$")
plan2019$Cuadro = stringr::str_remove(tolower(plan2019$Cuadro),".$")
plan2020$Nombre.variable = stringr::str_remove(plan2020$Nombre.variable,".$")

plan2020 = plan2020 %>% left_join(llave[,c("Nombre.FULL.2020.Reducida","Nombre.FULL.2019")],by = c("Nombre.variable" =  "Nombre.FULL.2020.Reducida")) %>% 
  left_join(plan2019[,c(5:8)], by = c("Nombre.FULL.2019"= "Nombre.variable")) %>%  select(- Nombre.FULL.2019, -Tipo.de.indicador)

plan2020$filtro[plan2020$Nombre.variable == "J5_1_1"] = "F1_1_1 ==1"

#### corregir regiones ####

enusc2020$enc_regionFIX = NA

enusc2020$enc_regionFIX[enusc2020$enc_region == 1] = "Tarapacá"
enusc2020$enc_regionFIX[enusc2020$enc_region == 2] = "Antofagasta"
enusc2020$enc_regionFIX[enusc2020$enc_region == 3] = "Atacama"
enusc2020$enc_regionFIX[enusc2020$enc_region == 4] = "Coquimbo"  
enusc2020$enc_regionFIX[enusc2020$enc_region == 5] = "Valparaíso" 
enusc2020$enc_regionFIX[enusc2020$enc_region == 6] = "O´Higgins"
enusc2020$enc_regionFIX[enusc2020$enc_region == 7] = "Maule"
enusc2020$enc_regionFIX[enusc2020$enc_region == 8] = "Biobío" 
enusc2020$enc_regionFIX[enusc2020$enc_region == 9] = "La Araucanía"
enusc2020$enc_regionFIX[enusc2020$enc_region == 10] =  "Los Lagos"
enusc2020$enc_regionFIX[enusc2020$enc_region == 11] =  "Aysen"
enusc2020$enc_regionFIX[enusc2020$enc_region == 12] =  "Magallanes"
enusc2020$enc_regionFIX[enusc2020$enc_region == 13] =  "Metropolitana"
enusc2020$enc_regionFIX[enusc2020$enc_region == 14] =  "Los Ríos"
enusc2020$enc_regionFIX[enusc2020$enc_region == 15] =  "Arica y Parinacota"
enusc2020$enc_regionFIX[enusc2020$enc_region == 16] =  "Ñuble"

#trimws(stringr::str_remove_all(rownames(data.frame(attr(enusc2020$enc_region,"labels"))),"Región de|la"))
enusc2020$enc_regionFIX = factor(enusc2020$enc_regionFIX,levels = c("Arica y Parinacota", "Tarapacá","Antofagasta","Atacama","Coquimbo",
                                                                    "Valparaíso","Metropolitana","O´Higgins","Maule","Ñuble", "Biobío",
                                                                    "La Araucanía","Los Ríos","Los Lagos","Aysen","Magallanes"))
## filtramos por kish ####
enusc2020 = enusc2020[enusc2020$Kish == 1,]


for(i in 1:55){ variable = plan2020$Nombre.variable[i]  
print(attr(enusc2020[[variable]],"labels")) 
}

enusc2020[[plan2020$Nombre.variable[5]]]

#### pasamos a dummyes según especicaciones de variables
# for(i in 1:55){
#   
#   variable =  plan2020$Nombre.variable[i] 
#   
#   if(any(grepl("Siempre",names(attr(enusc2020[[variable]],"labels"))))){
#     enusc2020[[variable]] = as.numeric(dplyr::if_else(enusc2020[[variable]] == 4,1,0))
#   }
#   
#   if(any(grepl("Aumentó",names(attr(enusc2020[[variable]],"labels"))))){
#     enusc2020[[variable]] = as.numeric(dplyr::if_else(enusc2020[[variable]] == 1,1,0))
#   }
#   
#   if(any(grepl("Sí",names(attr(enusc2020[[variable]],"labels"))))){
#     enusc2020[[variable]] = as.numeric(dplyr::if_else(enusc2020[[variable]] == 1,1,0))
#   }
#   
#   if(any(grepl("Si",names(attr(enusc2020[[variable]],"labels"))))){
#     enusc2020[[variable]] = as.numeric(dplyr::if_else(enusc2020[[variable]] == 1,1,0))
#   }
#   
#   print(table(enusc2020[,plan2020$Nombre.variable[i]]))
# }

## esta variable no está en el plan de tabulados
enusc2020$F1_1_1 = if_else(enusc2020$F1_1_1 == 1,1,0)


#### declaramos diseño complejo ####
dcPERS = svydesign(ids =~Conglomerado, strata = ~VarStrat, weights = ~Fact_Pers ,data = enusc2020)
dcHOG = svydesign(ids =~Conglomerado, strata = ~VarStrat, weights = ~Fact_Hog ,data = enusc2020)
options(survey.lonely.psu = "certainty")

extract_cat = function(var, total = T){
  #### condicional si se quiere con total #
  if(total == T){
    
    fe = list()
    ### si no tiene etiquetas  o si
    if(is.null(labelled::val_labels(var))){
      
      fe[[1]] <- sort(unique(var))
      
    }else{
      
      fe[[1]] = labelled::val_labels(var)
    }
    
    names(fe[[1]]) = "Total"
    
    for (i in 1:length(unique(var))){
      ### condicional si no tiene etiquetas  o si
      if(!is.null(names(labelled::val_labels(var)))){
        
        labelled::val_labels(var)[labelled::val_labels(var) %in% unique(var)] -> fi
        names(fi) <- NULL
        fe[[i+1]] <- sort(fi, decreasing = T)[i] 
        names(fe[[i+1]]) <- rev(names(labelled::val_labels(var)[labelled::val_labels(var) %in% unique(var)]))[i] 
        
      }else{
        
        ####  condicional si no tiene niveles (leves)  o si
        if(is.null(levels(var))){
          
          fe[[i+1]] <-  sort(unique(var), decreasing = T)[i]
          
          names(fe[[i+1]]) <- sort(unique(var), decreasing = T)[i]
          
        }else{
          
          fe[[i+1]] <- ordered(levels(var))[i]
          
          names(fe[[i+1]]) <- ordered(levels(var))[i]
          
        }
        
      }
    }
    ##### sin total
  }else{
    fe = list()
    
    for (i in 1:length(unique(var))){
      
      if(!is.null(names(labelled::val_labels(var)))){
        
        labelled::val_labels(var)[labelled::val_labels(var) %in% unique(var)] -> fi
        names(fi) <- NULL
        
        fe[[i]] <- sort(fi, decreasing = T)[i] 
        names(fe[[i]]) <-rev(names(labelled::val_labels(var)[labelled::val_labels(var) %in% unique(var)]))[i] }else{
          
          
          if(is.null(levels(var))){
            
            fe[[i]] <-  sort(unique(var), decreasing = T)[i]
            
            names(fe[[i]]) <- sort(unique(var), decreasing = T)[i]
            
          }else{
            
            fe[[i]] <- ordered(levels(var))[i]
            
            names(fe[[i]]) <- ordered(levels(var))[i]
            
          }
          
        }
    }
  }
  
  cat_names = names(unlist(fe))[!is.na(names(unlist(fe)))]
  
  cat = unlist(fe)[!is.na(names(unlist(fe)))]
  
  names(cat) =  NULL
  
  if(total == T){
    cat[1] = length(cat)
  }
  
  fe = list(catYname = fe, cat = cat, names = cat_names)
  
  return(fe)
  
}

replace_cat <- function(dominio, tabulado){
  
  cats = extract_cat(design$variables[[dominio]], total = F)
  
  tcat <-data.frame(Nombres = cats$names, valores = cats$cat)
  
  names(tcat)[2] <- dominio
  
  tcat[[2]] <- as.character(tcat[[2]])
  
  tabla <- left_join(tabulado, tcat) %>% select(last_col(), everything()) %>%  select(-dominio)
  
  names(tabla)[names(tabla) == "Nombres"] <- dominio
  
  tabla
  
}


calculo = "prop"
var=  "VP_DC"
dom = "rph_sexo+enc_region"
design = dcPERS
categorias_relevantes = "Sí"#2 #list("Sí" = 1)

class(2)
class("si")

creat_enusc <- function(var, design, dom, subpop = NULL,categorias_relevantes = NULL ){
calculo <- "prop"
#### creamos variables dumies ####  
  if(!is.null(categorias_relevantes)){
  
  if(is.character(categorias_relevantes)){  
    labels_var <- names(attr(design$variables[[var]],"labels"))
    dum_val <- stringr::str_detect(names(categorias_relevantes),labels_var)
    if(isFALSE(any(dum_val))){stop(paste0("the label \"",names(categorias_relevantes),"\" is not present in the variable ",var))}
    design$variables[[var]] = as.numeric(dplyr::if_else(design$variables[[var]] == categorias_relevantes[[1]],1,0))
  }else{
    labels_var <- unique(design$variables[[var]])
    dum_val <- (categorias_relevantes %in% labels_var)
    if(isFALSE(any(dum_val))){stop(paste0("the value \"",categorias_relevantes,"\" is not present in the variable ",var))}
    design$variables[[var]] = as.numeric(dplyr::if_else(design$variables[[var]] == categorias_relevantes[[1]],1,0))
    }

  }

#### funciones calidad ####
  
  funciones_cal = list(calidad::create_mean, calidad::create_prop,
                       calidad::create_tot_con, calidad::create_tot, calidad::create_median)
  funciones_eval = list(calidad::evaluate_mean, calidad::evaluate_prop, 
                        calidad::evaluate_tot_con, calidad::evaluate_tot, calidad::evaluate_median)
  
  if(calculo %in% "mean") {
    tipo = "mean"
    numCalc = 1
  }else if(calculo %in% "prop"){
    tipo = "objetivo"
    numCalc = 2
  }else if(calculo %in% "sum"){
    tipo = "total"
    numCalc = 3
  }else if(calculo %in% "count"){
    tipo = "total"
    numCalc = 4
  }else if(calculo %in% "median"){
    tipo = "median"
    numCalc = 5
  }    

test_dom <-trimws(unlist(stringr::str_extract_all(dom,"\\+")))
domD <-trimws(unlist(stringr::str_split(dom,"\\+")))

### alerta que solo se aceptan 2 niveles de desagregación
if(length(test_dom)>1){stop("Only accepts 2 disaggregation variables")}

### tablas con 1 nivel de desagregación ####
if(!grepl("\\+",dom)){

tprueba <- funciones_eval[[numCalc]](funciones_cal[[numCalc]](var,dominios = dom, disenio = design,subpop ,standard_eval = T), publicar = T)  %>% select(1:tipo,calidad) %>% mutate(calidad = ifelse(calidad == "fiable"," ",ifelse(calidad == "poco fiable","1","2")))

tprueba <- replace_cat(dom, tabulado = tprueba)

names(tprueba)[names(tprueba)== "calidad"] <- "Notas"  
names(tprueba)[names(tprueba)== tipo] <- "Total"

tprueba
### tablas con 2 nivel de desagregación ####
  }else{
    ### aseguramos que la variable con mas categorías quede hacia abajo
    if(nrow(unique(design$variables[domD[1]])) < nrow(unique(design$variables[domD[2]]))){
      dom3 <- domD[1]
      domD[1] <- domD[2]
      domD[2] <- dom3
    }
    
    tD1prueba <- funciones_eval[[numCalc]](funciones_cal[[numCalc]](var,dominios = domD[1], disenio = design,subpop , standard_eval = T), publicar = T)  %>% select(1:tipo,calidad)
    tD2prueba <- funciones_eval[[numCalc]](funciones_cal[[numCalc]](var,dominios = dom, disenio = design, subpop, standard_eval = T), publicar = T)  %>% select(1:tipo,calidad)
    
   tD1prueba <- replace_cat(domD[1],tD1prueba)
   tD2prueba <- replace_cat(domD[1],tD2prueba)
   tD2prueba <- replace_cat(domD[2],tD2prueba)

   orden <- c(domD[1],c(rbind(paste0(tipo,"_",extract_cat(design$variables[[domD[2]]], total = F)$name),paste0("calidad","_",extract_cat(design$variables[[domD[2]]], total = F)$name))))
   
   tD2prueba <- tD2prueba  %>% mutate(calidad = ifelse(calidad == "fiable"," ",ifelse(calidad == "poco fiable","1","2"))) %>% 
     pivot_wider(names_from = domD[2], values_from = c(tipo, calidad)) 
   
   nombre <- names(tD2prueba)
   
      orden <- orden[orden %in% nombre]
   
      tD2prueba <- tD2prueba %>% select(orden)
   
      num <- grep("calidad",names(tD2prueba))
      names(tD2prueba)[num] <- paste0("Notas",seq(length(num)))
    
      names(tD2prueba) <- names(tD2prueba) %>% stringr::str_remove_all(paste0(tipo,"_"))
    
      tD3prueba <-left_join(tD1prueba %>% mutate(calidad = ifelse(calidad == "fiable"," ",ifelse(calidad == "poco fiable","1","2"))), tD2prueba)
  
      names(tD3prueba)[names(tD3prueba)== "calidad"] <- "Notas"  
      names(tD3prueba)[names(tD3prueba)== tipo] <- "Total"
 
      names(tD3prueba)[grep("Notas",names(tD3prueba))] <- "Notas"
  
    tD3prueba   
 
      }

}

dcPERS = svydesign(ids =~Conglomerado, strata = ~VarStrat, weights = ~Fact_Pers ,data = enusc2020)

#rm(list = list(calculo, var, dom, design, categorias_relevantes))
rm(categorias_relevantes)

creat_enusc(var = "VP_DC", design = dcPERS, dom = "enc_region+rph_sexo",categorias_relevantes = list("Sí" = 1))
creat_enusc(var = "VP_DC", design = dcPERS, dom = "enc_region+rph_sexo",categorias_relevantes = list("SI" = 1))

creat_enusc(var = "VP_DC", design = dcPERS, dom = "enc_region+rph_sexo",categorias_relevantes = 1)
creat_enusc(var = "VP_DC", design = dcPERS, dom = "enc_region+rph_sexo",categorias_relevantes = 2)

creat_enusc(var = "VP_DC", design = dcPERS, dom = "enc_region+rph_sexo")

creat_enusc(var = "VP_DC", design = dcPERS, dom = "rph_sexo+enc_region")
creat_enusc(var = "VP_DC", design = dcPERS, dom = "enc_region+rph_sexo")
creat_enusc(var = "VA_DC", design = dcHOG, dom = "enc_region+rph_sexo")


enusc2020 = read_sav("base_usuario_ENUSC2020.sav")
plan_ejemplo = read.xlsx("plan_tabulado_ejemplo.xlsx")
create_enusc_tabulated_plan
design = dcPERS

tabulated_plan = plan_ejemplo
categorias_relevantes = list("Siempre" = 4, "Aumentó" = 1, "Sí" = 1, "Si" = 1)

############################################  
  
labels_var <- names(attr(design$variables[[variable]],"labels"))

dum_val <- categorias_relevantes[stringr::str_detect(cat_rel,labels_var)]

design$variables[[variable]] = as.numeric(dplyr::if_else(design$variables[[variable]] == dum_val[[1]],1,0))




creat_enusc(var = "VA_DC", design = dcHOG, dom = "enc_region+rph_sexo")










