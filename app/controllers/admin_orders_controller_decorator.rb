# -*- coding: utf-8 -*-
require "spreadsheet"
require 'stringio'
Admin::OrdersController.class_eval do
  def post_blank_1
    @order = Order.find_by_number(params[:id])
    nalozhenniy = @order.adjustment_total > 0 ? @order.total.to_i : nil

    Spreadsheet.client_encoding = 'UTF-8'
    book = Spreadsheet.open File.join(Rails.root, 'public/russian_post/pered.xls')
    sheet = book.worksheet 0
    address = @order.ship_address
    full_name = "#{address.lastname} #{address.firstname} #{address.secondname}"
    oblast = address.state.nil? ? '' : address.state.name 
    gorod = address.city
    adres_v_gorode = address.address1
    zip = address.zipcode
    
    sheet[51,10] = full_name
    sheet[56,11] = oblast + ', ' + gorod
    sheet[60,4] = adres_v_gorode
    sheet[65,45] = zip.mb_chars[0].to_s
    sheet[65,48] = zip.mb_chars[1].to_s
    sheet[65,51] = zip.mb_chars[2].to_s
    sheet[65,54] = zip.mb_chars[3].to_s
    sheet[65,57] = zip.mb_chars[4].to_s
    sheet[65,60] = zip.mb_chars[5].to_s
    
    sheet[70,13] = 'Антипин Константин Валерьевич'
    sheet[74,10] = 'г. Москва'
    sheet[78,4] = 'до востребования'
    # sheet[78,4] = '3й проезд Марьиной Рощи д.5 кв 124'
    sheet[78,71] = '1'
    sheet[78,74] = '2'
    sheet[78,77] = '9'
    sheet[78,80] = '5'
    sheet[78,83] = '9'
    sheet[78,86] = '4'

    sheet[131,10] = full_name
    sheet[136,12] = oblast + ', ' + gorod
    sheet[142,4] = adres_v_gorode
    sheet[142,71] = zip.mb_chars[0].to_s
    sheet[142,74] = zip.mb_chars[1].to_s
    sheet[142,77] = zip.mb_chars[2].to_s
    sheet[142,80] = zip.mb_chars[3].to_s
    sheet[142,83] = zip.mb_chars[4].to_s
    sheet[142,86] = zip.mb_chars[5].to_s

    sheet[92,121] = full_name
    sheet[100,131] = oblast + ', ' + gorod
    sheet[104,111] = adres_v_gorode

    # Если строка больше 42 букв, то делаем перенос на следующую
    # строку. отменено.
    # https://www.google.com/url?url=http://ru.wikibooks.org/wiki/Ruby/%25D0%259F%25D0%25BE%25D0%25B4%25D1%2580%25D0%25BE%25D0%25B1%25D0%25BD%25D0%25B5%25D0%25B5_%25D0%25BE_%25D1%2581%25D1%2582%25D1%2580%25D0%25BE%25D0%25BA%25D0%25B0%25D1%2585%23.D0.9F.D0.B5.D1.80.D0.B5.D0.BD.D0.BE.D1.81_.D0.BF.D0.BE_.D1.81.D0.BB.D0.BE.D0.B2.D0.B0.D0.BC&rct=j&q=ruby+%D0%BF%D0%B5%D1%80%D0%B5%D0%BD%D0%BE%D1%81+%D1%81%D1%82%D1%80%D0%BE%D0%BA%D0%B8&usg=AFQjCNE8Ln36ay94ywERjgM8frRuPGXUuw&sa=X&ei=gZcST-exK4Gl-gaH75TvAg&ved=0CCwQygQwAA
    # propisju_string = RuPropisju.rublej(@order.total).gsub(/(.{1,42})( +|$\n?)|(.{1,42})/, "\\1\\3\n").split(/\n/)
    if @order.outstanding_balance > 0 and nalozhenniy
      sheet[41,4] = RuPropisju.rublej(@order.total) # объявленная ценность
      sheet[46,4] = RuPropisju.rublej(@order.total) #наложенный платёж
      sheet[34,111] = RuPropisju.rublej(@order.total) # почтовый перевод
      sheet[126, 14] = nalozhenniy
      sheet[126, 46] = nalozhenniy
      sheet[29, 163] = nalozhenniy
    end
    # Если стоимость посылки не была рассчитана (труднодоступный регион),
    # то написать в уголочке стоимость для операторов
    sheet[149,0] = @order.item_total if (@order.outstanding_balance > 0 and not nalozhenniy)
    unless @order.outstanding_balance > 0
      sheet[149,0] = 'опл' # слева в уголке
      sheet[41,4] = 'один рубль' # объявленная ценность
    end

    sheet[0,0] = ''

    data = StringIO.new ''
    book.write data
    send_data data.string, :type=>"application/excel", :disposition=>'attachment', :filename => "Форма_наложенного_лицевая_#{@order.number}.xls"

    # book.write File.join(Rails.root, 'public/russian_post/pered_out.xls')
    # send_file File.join(Rails.root, 'public/russian_post/pered_out.xls'), :type=>"application/xls" 
    # redirect_to :back
  end
  def post_blank_2
    @order = Order.find_by_number(params[:id])
    Spreadsheet.client_encoding = 'UTF-8'
    book = Spreadsheet.open File.join(Rails.root, 'public/russian_post/zad.xls')
    sheet = book.worksheet 0

    address = @order.ship_address
    sheet[37,11] = address.lastname
    sheet[41,3] = "#{address.firstname} #{address.secondname}"

    sheet[0,0] = ''

    data = StringIO.new ''
    book.write data
    send_data data.string, :type=>"application/excel", :disposition=>'attachment', :filename => "Форма_наложенного_обратная_#{@order.number}.xls"

    # book.write File.join(Rails.root, 'public/russian_post/pered_out.xls')
    # send_file File.join(Rails.root, 'public/russian_post/pered_out.xls'), :type=>"application/xls" 
    # redirect_to :back
  end

end

