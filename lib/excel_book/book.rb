# encoding: utf-8

#
# Excelを操作するためのGemのwin32oleを自分が使いやすくカスタマイズしたもの
#
module ExcelBook
  class Book
    @@open_books = []

    attr_reader :book

    class << Book
      # Excelのファイルを開く。ブロックが渡されていたら明示的にcloseしなくても良い。
      # @param filepath オープンするファイルのパス
      def open(filepath, options = {mode: 'r'}, &block)
        if File.basename(filepath) == filepath
          filepath = File.expand_path(filepath, Dir::pwd)
        end

        new(filepath, options, &block)
      end

      # シートのコピーを行う
      # @param from original sheet
      # @param dest target sheet
      # @param force force copy
      def copy(from, dest, force: false)

        display_alerts = @@excel.DisplayAlerts
        @@excel.DisplayAlerts = false
      
        # 強制的に置き換える場合は、一度ファイル名を変えてからコピーし、その後削除を行う。
        if force && from.name == dest.name
          dest.name = 'ranDom@String'
          from.sheet.copy(dest.sheet)
          Book.delete(dest)
        else
          from.sheet.Copy(dest.sheet)
        end

        @@excel.DisplayAlerts = display_alerts
      end

      # シートの削除を行う
      # @param sheet delete target sheet.
      def delete(sheet)
        display_alerts = @@excel.DisplayAlerts
        @@excel.DisplayAlerts = false
        sheet.delete
        @@excel.DisplayAlerts = display_alerts
      end

      #
      # 現在開いているExcel Bookの一覧を返す
      #
      def active_books
        @@open_books.values
      end
    end

    #
    # 渡されたファイル名からbookインスタンスを生成。
    #   mode r : 読み取りモード
    #   mode w : 書き込みモード
    #
    def initialize(filepath, options = {}, &block)
      @@excel ||= WIN32OLE.new('Excel.Application')
      fso = WIN32OLE.new('Scripting.FileSystemObject')
      filepath.encode!('UTF-8')

      case options[:mode]
      when 'r'
        begin
          raise Errno::ENOENT unless File.exist?(filepath)
          @book = excel.Workbooks.Open(fso.GetAbsolutePathName(filepath))
        rescue => ex
          quit
          puts "#{ex.message}"
        end
      when 'w'
        begin
          if File.exist?(filepath)
            @book = excel.Workbooks.Open(fso.GetAbsolutePathName(filepath))
          else
            filepath.gsub!('/', '\\')
            excel.Workbooks.Add.saveAs(filepath)
            @book = excel.Workbooks.Open(fso.GetAbsolutePathName(filepath))
          end
        rescue => ex
          quit
          puts ex.message
        end
      else
        @book = excel.Workbooks.Open(fso.GetAbsolutePathName(filepath))
      end

      @@open_books[@book.object_id] = @book

      if block
        begin
          yield(self)
        ensure
          close
          quit
        end
      else
        book
      end
    end

    def excel
      @@excel
    end

    #
    # シートの追加を行う。
    # [param]
    #   シートの名前
    #
    def add_sheet(param = nil)
      display_alerts = excel.DisplayAlerts
      excel.DisplayAlerts = false

      if param.instance_of?(String)
        sheet = book.Worksheets.Add({ after: last.sheet })
        sheet.name = param if param
      elsif param.instance_of?(Sheet)
        param.sheet.Copy({ after: last.sheet })
        new_sheet  = last
      else
        raise '予期せぬパラメータです。'
      end
      excel.DisplayAlerts = display_alerts
        
      Sheet.new(sheet)
    end

    #
    # シートの名前一覧を取得する。
    #
    def sheet_names
      names = []
    
      book.Worksheets.each do |sheet|
        names << sheet.name
      end

      names
    end


    #
    # シートの一覧を取得する
    # [param] 
    #   シート番号
    #
    # 返り値 : シートの配列
    #
    def sheets(param = nil)
      return Sheet.new(book.Worksheets(param)) if param

      sheets = []

      book.Worksheets.each do |sheet|
        sheets << Sheet.new(sheet)
      end

      sheets
    end

    #
    # シートの一番最初を取り出す
    #
    def first
      Sheet.new(book.Worksheets(1))
    end

    #
    # 二番目のシートを取り出す
    #
    def second
      Sheet.new(book.Worksheets(2))
    end

    #
    # 一番最後のシートを取り出す
    #
    def last
      Sheet.new(book.Worksheets(book.Worksheets.Count))
    end

    #
    # シートの削除
    # [param]
    #   シートのインデックス or シートの名前
    #
    def delete(param)
      display_alerts = excel.DisplayAlerts
      excel.DisplayAlerts = false
      book.Worksheets(param).Delete
      excel.DisplayAlerts = display_alerts
    end

    # 
    # 保存を行う
    #
    def save
      display_alerts = excel.DisplayAlerts
      excel.DisplayAlerts = false
      book.save
      excel.DisplayAlerts = display_alerts
    end

    #
    # シートを閉じる
    #
    def close
      @@open_books.delete(book.object_id)
      book.close
      quit
    end

    #
    # Excelを閉じる、アクティブなBookがある場合は終了しない。
    #
    def quit
      excel.Quit if @@open_books.empty?
    end
  end
end