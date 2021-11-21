pacman::p_load(
  data.table,
  tidyverse,
  rvg,
  xlsx,
  officer
)

plot_gg <- 
  iris %>% 
  ggplot() +
  aes(Sepal.Length,Sepal.Width, color = Species) +
  geom_point() +
  theme_bw() +
  labs(
    title = "Title",
    subtitle = "SubTitle",
    color = ""
  ) +
  theme(
    legend.position = "top"
  )

text_dt <- 
  as.data.table(
    read.xlsx(
      here::here("text_files", "text_for_pptx.xlsx"),
      1
    )
  )

text_dt$page_number <- as.character(text_dt$page_number)
text_dt$analysis_number <- as.character(text_dt$analysis_number)


#--------------------#

# load temple ----
my_pres <- 
  read_pptx(here::here("pptx_templates", "final_template.pptx")) %>% 
  remove_slide(index = 1)

# add slide from template function ----
add_sld <- function(pers_obj, layout_name, plot_name = NULL, page_num = NULL){
  if(layout_name == "cover_page"){
    return(
      add_slide(pers_obj, layout = layout_name, master = "Office Theme")
    )
  } else if(layout_name == "chapter_open"){
    return(
      add_slide(pers_obj, layout = layout_name, master = "Office Theme") %>% 
        ph_with(value = "123", location = ph_location_label(ph_label = "num_ovdim")) %>% 
        ph_with(value = "234", location = ph_location_label(ph_label = "num_misra")) %>% 
        ph_with(value = "345", location = ph_location_label(ph_label = "sal_mean")) %>% 
        ph_with(value = "456", location = ph_location_label(ph_label = "sal_mid"))
    )
  } else if(layout_name == "p_0"){
    return(
      add_slide(pers_obj, layout = layout_name, master = "Office Theme") %>% 
        ph_with(value = text_dt[page_number == page_num, analysis_number], location = ph_location_label(ph_label = "analysis_number")) %>%  
        ph_with(value = text_dt[page_number == page_num, title_text], location = ph_location_label(ph_label = "title_text")) %>% 
        ph_with(value = text_dt[page_number == page_num, main_text], location = ph_location_label(ph_label = "main_text")) %>% 
        ph_with(value = text_dt[page_number == page_num, page_number], location = ph_location_label(ph_label = "page_number"))
    )
  } else if(layout_name %in% c("p_50", "p_75", "p_90")){
    return(
      add_slide(pers_obj, layout = layout_name, master = "Office Theme") %>% 
        ph_with(value = text_dt[page_number == page_num, analysis_number], location = ph_location_label(ph_label = "analysis_number")) %>%  
        ph_with(value = text_dt[page_number == page_num, title_text], location = ph_location_label(ph_label = "title_text")) %>% 
        ph_with(value = text_dt[page_number == page_num, main_text], location = ph_location_label(ph_label = "main_text")) %>% 
        ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location = ph_location_label(ph_label = "plot_place")) %>% 
        ph_with(value = text_dt[page_number == page_num, page_number], location = ph_location_label(ph_label = "page_number"))
    )
  } else if(layout_name == "p_100"){
    return(
      add_slide(pers_obj, layout = layout_name, master = "Office Theme") %>% 
        ph_with(value = text_dt[page_number == page_num, analysis_number], location = ph_location_label(ph_label = "analysis_number")) %>%  
        ph_with(value = text_dt[page_number == page_num, title_text], location = ph_location_label(ph_label = "title_text")) %>% 
        ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location = ph_location_label(ph_label = "plot_place")) %>% 
        ph_with(value = text_dt[page_number == page_num, page_number], location = ph_location_label(ph_label = "page_number"))
    )
  }
}
    
#--------------------#

# bulid the output ----
my_pres <- 
  add_sld(my_pres, "cover_page") %>% 
  add_sld("chapter_open") %>% 
  add_sld("p_0",
          plot_name = plot_gg,
          page_num = "1") %>% 
  add_sld("p_50",
          plot_name = plot_gg,
          page_num = "5") %>% 
  add_sld("p_75",
          plot_name = plot_gg,
          page_num = "4") %>% 
  add_sld("p_90",
          plot_name = plot_gg,
          page_num = "3") %>% 
  add_sld("p_100",
          plot_name = plot_gg,
          page_num = "2")


# save pptx ----
output_name <- "ophir_is_cool"
print(my_pres, here::here("outputs", glue::glue("{output_name}_{Sys.Date()}.pptx")))
