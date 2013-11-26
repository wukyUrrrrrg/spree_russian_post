class SpreeRussianPostHooks < Spree::ThemeSupport::HookListener
  # custom hooks go here
  insert_after :admin_order_show_buttons do
    %(<%= button_link_to(I18n::t("post_blank_side_1"), admin_order_post_blank_1_url(@order)) %>
    <%= button_link_to(I18n::t("post_blank_side_2"), admin_order_post_blank_2_url(@order)) %>)
  end

end
