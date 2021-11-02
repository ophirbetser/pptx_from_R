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
      "/Users/ophirbetser/Ophir/R PROJECTS/pptx_from_R/text_for_pptx.xlsx",
      1
    )
  )

#--------------------#

# load temple ----
my_pres <- 
  read_pptx(here::here("pptx_tempalte_for_heb.pptx")) %>% 
  remove_slide(index = 1)

# add slide from template function ----
add_sld <- function(pers_obj, layout_name, plot_name = NULL, page_num = NULL){
  if(layout_name %in% c("front_page", "intro_page", "empty_page")){
       return(
         add_slide(pers_obj, layout = layout_name, master = "Office Theme")
       )
  } else if(layout_name %in% c("layout_full_plot")){
    return(
      add_slide(pers_obj, layout = "layout_full_plot", master = "Office Theme") %>% 
        ph_with(value = text_dt[page == page_num, header], location = ph_location_label(ph_label = "num")) %>%  
        ph_with(value = text_dt[page == page_num, title], location = ph_location_label(ph_label = "text")) %>% 
        ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location =  ph_location_label(ph_label = "plot")) %>% 
        ph_with(value = text_dt[page == page_num, page], location = ph_location_label(ph_label = "page_number"))
    )
  } else if(layout_name %in% c("layout_half", "layout_75")){
    return(
      add_slide(pers_obj, layout = layout_name, master = "Office Theme") %>% 
        ph_with(value = text_dt[page == page_num, header], location = ph_location_label(ph_label = "num")) %>%  
        ph_with(value = text_dt[page == page_num, text], location = ph_location_label(ph_label = "text")) %>% 
        ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location = ph_location_label(ph_label = "plot")) %>% 
        ph_with(value = text_dt[page == page_num, page], location = ph_location_label(ph_label = "page_number"))
    )
  } else if(layout_name %in% c("layout_full_plot_with_text")){
    return(
      add_slide(pers_obj, layout = layout_name, master = "Office Theme") %>% 
        ph_with(value = text_dt[page == page_num, header], location = ph_location_label(ph_label = "num")) %>%  
        ph_with(value = text_dt[page == page_num, title], location = ph_location_label(ph_label = "title_text")) %>% 
        ph_with(value = text_dt[page == page_num, text], location = ph_location_label(ph_label = "text")) %>% 
        ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location = ph_location_label(ph_label = "plot")) %>% 
        ph_with(value = text_dt[page == page_num, page], location = ph_location_label(ph_label = "page_number")) 
    )
  } else if(layout_name %in% c("layout_full_text")){
    return(
      add_slide(pers_obj, layout = layout_name, master = "Office Theme") %>% 
        ph_with(value = text_dt[page == page_num, header], location = ph_location_label(ph_label = "num")) %>%  
        ph_with(value = text_dt[page == page_num, title], location = ph_location_label(ph_label = "header")) %>% 
        ph_with(value = text_dt[page == page_num, text], location = ph_location_label(ph_label = "text")) %>% 
        ph_with(value = text_dt[page == page_num, page], location = ph_location_label(ph_label = "page_number"))
    )
  }
}

# bulid the output
my_pres <- 
  add_sld(my_pres, "front_page") %>% 
  add_sld("intro_page") %>% 
  add_sld("layout_full_plot",
          plot_name = plot_gg,
          page_num = "1") %>% 
  add_sld("layout_half",
          plot_name = plot_gg,
          page_num = "2") %>% 
  add_sld("layout_75",
          plot_name = plot_gg,
          page_num = "3") %>% 
  add_sld("layout_full_plot_with_text",
          plot_name = plot_gg,
          page_num = "4") %>% 
  add_sld("layout_full_text",
          page_num = "5") %>% 
  add_sld("layout_75",
          plot_name = plot_gg,
          page_num = "3") %>% 
  add_sld("layout_75",
          plot_name = plot_gg,
          page_num = "3") %>% 
  add_sld("layout_full_plot",
          plot_name = plot_gg,
          page_num = "1") %>% 
  add_sld("layout_half",
          plot_name = plot_gg,
          page_num = "2") %>% 
  add_sld("empty_page")
  

# save pptx ----
output_name <- "my_pes"
print(my_pres, here::here(glue::glue("{output_name}_{Sys.Date()}.pptx")))
