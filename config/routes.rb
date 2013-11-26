Rails.application.routes.draw do

  # Add your extension routes here
  namespace :admin do
    get 'orders/:id/post_blank_1' =>  "orders#post_blank_1", :as => :order_post_blank_1
    get 'orders/:id/post_blank_2' =>  "orders#post_blank_2", :as => :order_post_blank_2
  end

end
