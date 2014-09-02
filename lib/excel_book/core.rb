# encoding: utf-8

#
# Excelを操作するためのGemのwin32oleを自分が使いやすくカスタマイズしたもの
#
class ExcelBook
  @@open_books = []

  attr_reader :book

  class << ExcelBook

    #
    # Excelのファイルを開く。ブロックが渡されていたら明示的にcloseしなくても良い。
    # [filename]
    #   オープンするファイルのパス
    #
    def new(filename)
      book = super

      if block_given?

        begin
          yield(book)
        ensure
          book.close
          book.quit
        end
      else
        book
      end
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
    def copy(from, dest, force: true)

      display_alerts = @@excel.DisplayAlerts
      @@excel.DisplayAlerts = false
      
      # 強制的に置き換える場合は、一度ファイル名を変えてからコピーし、その後削除を行う。
      if force
        dest.name = 'randomStrign'
        from.copy(dest)
        ExcelBook.delete(dest)
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

    alias_method :open, :new
  end

  #
  # 渡されたファイル名からbookインスタンスを生成。
  #
  def initialize(filename)
    @@excel ||= WIN32OLE.new('Excel.Application')
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    @book = excel.Workbooks.Open(fso.GetAbsolutePathName(filename))
    @@open_books[@book.object_id] = @book
  end

  def excel
    @@excel
  end

  #
  # シートの名前一覧を取得する。
  #
  def sheet_names
    names = []
    
    @book.Worksheets.each do |sheet|
      names << sheet.name
    end

    names
  end


  #
  # シートの一覧を取得する
  # [sheet_num] 
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
  end

  #
  # Excelを閉じる、アクティブなBookがある場合は終了しない。
  #
  def quit
    excel.Quit if @@open_books.empty?
  end
end