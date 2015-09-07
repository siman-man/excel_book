require 'spec_helper'

describe ExcelBook do
  it 'ExcelBook has a version.' do
    expect(ExcelBook::VERSION).not_to be nil
  end

  context 'Create Excel Book' do
    it 'ExcelBook can create new excel file.' do
      filepath = File.expand_path('example.xlsx', TEMP_DIR)

      FileUtils.rm_f(filepath) if File.exist?(filepath)

      book = ExcelBook::Book.open(filepath, mode: 'w')
      book.close

      expect(File.exist?(filepath)).to be_truthy
    end

    it 'ExcelBook can create new excel file(Block).' do
      filepath = File.expand_path('example_block.xlsx', TEMP_DIR)

      FileUtils.rm_f(filepath) if File.exist?(filepath)

      ExcelBook::Book.open(filepath, mode: 'w') do |book|
      end

      expect(File.exist?(filepath)).to be_truthy
    end
  end

  context 'Read Cell Value' do
    it 'ExcelBook can read cell value' do
      filepath = File.expand_path('sample.xlsx', TEMP_DIR)

      begin
        book = ExcelBook::Book.open(filepath)
        sheet = book.first
        a1 = sheet['A1']
        b1 = sheet['B1']
      ensure
        book.close
      end

      expect(a1).to eq 1.0
      expect(b1).to eq 'A'
    end
  end
end
