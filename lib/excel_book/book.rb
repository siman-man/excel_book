# encoding: utf-8

#
# Excelを操作するためのGemのwin32oleを自分が使いやすくカスタマイズしたもの
#
module ExcelBook
  class Book
    @@open_books = []

    attr_reader :book

    class << Book

      #
      # Excelのファイルを開く。ブロックが渡されていたら明示的にcloseしなくても良い。
      # [filepath]
      #   オープンするファイルのパス
      #
      def open(filepath, options = {mode: 'r'}, &block)
        new(filepath, options, &block)
      end

      #
      # シートのコピーを行う
      # [from]
      #   コピー元シート
      # [dest]
      #   コピー先シート
      # [force]
      #   trueだと強制的に置き換える   
      #
      def copy(from, dest, force: false)

        display_alerts = @@excel.DisplayAlerts
        @@excel.DisplayAlerts = false
      
        # 強制的に置き換える場合は、一度ファイル名を変えてからコピーし、その後削除を行う。
        if force && from.name == dest.name
          dest.name = 'randomStrign'
          from.copy(dest)
          Book.delete(dest)
        else
          from.Copy(dest)
        end

        @@excel.DisplayAlerts = display_alerts
      end

      #
      # シートの削除を行う
      # [sheet]
      #   削除対象のシート
      #
      def delete(sheet)
        display_alerts = @@excel.DisplayAlerts
        @@excel.DisplayAlerts = false
        sheet.Delete
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
    #
    def initialize(filepath, options = {}, &block)
      @@excel ||= WIN32OLE.new('Excel.Application')
      fso = WIN32OLE.new('Scripting.FileSystemObject')

      case options[:mode]
      when 'r'
        begin
          raise Errno::ENOENT unless File.exist?(filepath)
          @book = excel.Workbooks.Open(fso.GetAbsolutePathName(filepath))
        rescue => ex
          quit
          puts "#{filepath} : #{ex.message}"
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
    #
    def add_sheet(name = nil)
      sheet = book.Worksheets.Add({ after: last })
      sheet.name = name if name
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
      return book.Worksheets(param) if param

      sheets = []

      book.Worksheets.each do |sheet|
        sheets << sheet
      end

      sheets
    end

    #
    # シートの一番最初を取り出す
    #
    def first
      book.Worksheets(1)
    end

    #
    # 二番目のシートを取り出す
    #
    def second
      book.Worksheets(2)
    end

    #
    # 一番最後のシートを取り出す
    #
    def last
      book.Worksheets(book.Worksheets.Count)
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
      book.save
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