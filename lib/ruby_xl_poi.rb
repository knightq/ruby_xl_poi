# frozen_string_literal: true

require 'rjb'
require 'forwardable'

require 'ruby_xl_poi/version'
require 'ruby_xl_poi/workbook'

require 'ruby_xl_poi/worksheet/sheet'
require 'ruby_xl_poi/worksheet/row'
require 'ruby_xl_poi/worksheet/cell'

module RubyXlPoi
  def self.init
    libs = Dir.glob("#{File.dirname(__FILE__)}/../vendor/java/apache/*.jar").join(':')
    Rjb.load(libs, ['-Xmx512M'])

    @cell_class = Rjb.import('org.apache.poi.ss.usermodel.Cell')

    Rjb.import('org.apache.poi.openxml4j.opc.OPCPackage')
    Rjb.import('org.apache.poi.ss.usermodel.Cell')
    Rjb.import('org.apache.poi.ss.usermodel.Row')
    Rjb.import('org.apache.poi.ss.usermodel.Sheet')
    Rjb.import('org.apache.poi.ss.usermodel.Workbook')
    Rjb.import('org.apache.poi.xssf.usermodel.XSSFCell')
    Rjb.import('org.apache.poi.xssf.usermodel.XSSFRow')
    Rjb.import('org.apache.poi.xssf.usermodel.XSSFSheet')
    Rjb.import('org.apache.poi.xssf.usermodel.XSSFWorkbook')

    Rjb.import('org.apache.poi.ss.usermodel.CreationHelper')
    Rjb.import('org.apache.poi.ss.usermodel.FormulaEvaluator')

    # You can import all java classes that you need
    @loaded = true
  end

  class << self
    attr_reader :cell_class
  end

  def self.load(file)
    init unless @loaded
    Workbook.load file
  end
end
