
t_v_short <- "וואלק וואלק"
t_long <- "וואלק וואלק"

plot_name <- 
  iris %>% ggplot() +
    aes(Sepal.Length,Sepal.Width) +
    geom_point()

my_pres <- 
  read_pptx(here::here("pptx_templates", "final_template.pptx")) %>% 
  remove_slide(index = 1) %>%  
  
  add_slide(layout = "cover_page", master = "Office Theme") %>% 
  
  add_slide(layout = "chapter_open", master = "Office Theme") %>% 
  ph_with(value = "123", location = ph_location_label(ph_label = "num_ovdim")) %>% 
  ph_with(value = "234", location = ph_location_label(ph_label = "num_misra")) %>% 
  ph_with(value = "345", location = ph_location_label(ph_label = "sal_mean")) %>% 
  ph_with(value = "456", location = ph_location_label(ph_label = "sal_mid")) %>% 
  
  add_slide(layout = "p_0", master = "Office Theme") %>% 
  ph_with(value = "4.1", location = ph_location_label(ph_label = "analysis_number")) %>% 
  ph_with(value = t_v_short, location = ph_location_label(ph_label = "title_text")) %>% 
  ph_with(value = t_long, location = ph_location_label(ph_label = "main_text")) %>% 
  ph_with(value = "12", location = ph_location_label(ph_label = "page_number")) %>% 
  
  add_slide(layout = "p_50", master = "Office Theme") %>% 
  ph_with(value = "4.2", location = ph_location_label(ph_label = "analysis_number")) %>% 
  ph_with(value = t_v_short, location = ph_location_label(ph_label = "title_text")) %>% 
  ph_with(value = t_long, location = ph_location_label(ph_label = "main_text")) %>% 
  ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location = ph_location_label(ph_label = "plot_place")) %>% 
  ph_with(value = "13", location = ph_location_label(ph_label = "page_number")) %>% 
  
  add_slide(layout = "p_75", master = "Office Theme") %>% 
  ph_with(value = "4.3", location = ph_location_label(ph_label = "analysis_number")) %>% 
  ph_with(value = t_v_short, location = ph_location_label(ph_label = "title_text")) %>% 
  ph_with(value = t_long, location = ph_location_label(ph_label = "main_text")) %>% 
  ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location = ph_location_label(ph_label = "plot_place")) %>% 
  ph_with(value = "14", location = ph_location_label(ph_label = "page_number")) %>% 
  
  add_slide(layout = "p_90", master = "Office Theme") %>% 
  ph_with(value = "4.4", location = ph_location_label(ph_label = "analysis_number")) %>% 
  ph_with(value = t_v_short, location = ph_location_label(ph_label = "title_text")) %>% 
  ph_with(value = t_long, location = ph_location_label(ph_label = "main_text")) %>% 
  ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location = ph_location_label(ph_label = "plot_place")) %>% 
  ph_with(value = "15", location = ph_location_label(ph_label = "page_number")) %>% 
  
  add_slide(layout = "p_100", master = "Office Theme") %>% 
  ph_with(value = "4.5", location = ph_location_label(ph_label = "analysis_number")) %>% 
  ph_with(value = t_v_short, location = ph_location_label(ph_label = "title_text")) %>% 
  ph_with(value = rvg::dml(ggobj = plot_name, bg = "transparent"), location = ph_location_label(ph_label = "plot_place")) %>% 
  ph_with(value = "16", location = ph_location_label(ph_label = "page_number"))


  
  
  
# save pptx ----
output_name <- "new_temp"
print(my_pres, here::here("outputs", glue::glue("{output_name}_{Sys.Date()}.pptx")))
