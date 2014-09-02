require 'spec_helper'

describe ExcelBook do
  it 'ExcelBook has a version.' do
    expect(ExcelBook::VERSION).not_to be nil
  end

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
