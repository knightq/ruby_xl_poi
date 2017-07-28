# frozen_string_literal: true

module RubyXlPoi
  class Workbook
    extend Forwardable

    attr_accessor :j_book, :sheets

    def self.load(file)
      @file_name = file

      @workbook_factory_class = Rjb.import('org.apache.poi.ss.usermodel.WorkbookFactory')
      @poifs_class = Rjb.import('org.apache.poi.poifs.filesystem.POIFSFileSystem')
      @opcp_class = Rjb.import('org.apache.poi.openxml4j.opc.OPCPackage')
      @xssf_class = Rjb.import('org.apache.poi.xssf.usermodel.XSSFWorkbook')
      @file_input_class = Rjb.import('java.io.FileInputStream')
      @file_input = @file_input_class.new(file)

      fs = if xlsm?(file)
             @opcp_class.open(@file_input)
           else
             @poifs_class.new(@file_input)
           end
      j_book = xlsm?(file) ? @xssf_class.new(fs) : @workbook_factory_class.create(fs)
      new(j_book)
    end

    def initialize(j_book = nil)
      self.j_book = j_book
      self.sheets = {}
      return unless j_book
      nr_of_sheets = j_book.getNumberOfSheets.to_i - 1
      return if nr_of_sheets.negative?
      (0..nr_of_sheets).each do |sheet_index|
        j_sheet = j_book.getSheetAt(sheet_index)
        sheet_name = j_book.getSheetName(sheet_index)
        sheets.store sheet_name, Worksheet::Sheet.new(self, j_sheet) if j_sheet
      end
    end

    def self.xlsm?(file_name)
      file_name =~ /\.xlsm$/
    end

    def create_sheet(name)
      sheets.store name, Worksheet::Sheet.new(self, j_book.createSheet(name))
    end

    def clone_sheet(index)
      sheets.store "Sheet #{index + 1}", Worksheet::Sheet.new(self, j_book.cloneSheet(index))
    end

    def remove_sheet_at(index)
      j_book.removeSheetAt(index)
      @sheets.delete_at(index)
    end

    # Get sheet by name
    def_delegator :sheets, :[], :[]

    def save(file_name = @file_name)
      @file_output_class ||= Rjb.import('java.io.FileOutputStream')
      out = @file_output_class.new(file_name)

      begin
        j_book.write(out)
      ensure
        out.close
      end
    end

    def _evaluator
      @_evaluator ||= j_book.getCreationHelper.createFormulaEvaluator
    end
  end
end
