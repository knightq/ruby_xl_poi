# frozen_string_literal: true

module RubyXlPoi
  module Worksheet
    class Cell
      attr_accessor :j_cell, :row, :index

      def initialize(row, j_cell = nil, index = 1)
        self.j_cell = j_cell
        self.row = row
        self.index = index
      end

      def type
        self.class.symbol_type(j_cell.getCellType)
      end

      def blank?
        type == :blank
      end

      def formula?
        type == :formula
      end

      def numeric?
        type == :numeric
      end

      def val
        # type = self.class.symbol_type(sheet.book._evaluator.evaluateFormulaCell(cell))
        type_pref = formula? ? :numeric : type
        return if type_pref == :blank
        j_cell.send("get#{type_pref.capitalize}CellValue")
      end

      def to_s
        val
      end

      def inspect
        "row: #{row&.index}, column: #{index}, type: #{type}, val: #{val}"
      end

      def val=(value)
        j_cell.setCellValue(value)
      end

      def change_contents(data, _formula = nil)
        j_cell.send('setCellValue', data)
      end

      def self.symbol_type(constant)
        types[constant]
      end

      def self.types
        @types ||= begin
          cell = ::RubyXlPoi.cell_class
          {
            cell.CELL_TYPE_BOOLEAN => :boolean,
            cell.CELL_TYPE_NUMERIC => :numeric,
            cell.CELL_TYPE_STRING => :string,
            cell.CELL_TYPE_BLANK => :blank,
            cell.CELL_TYPE_ERROR => :error,
            cell.CELL_TYPE_FORMULA => :formula
          }
        end
      end
      private_class_method :types
    end
  end
end
