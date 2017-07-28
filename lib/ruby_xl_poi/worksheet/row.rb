# frozen_string_literal: true

module RubyXlPoi
  module Worksheet
    class Row
      attr_accessor :cells, :index, :j_row, :sheet

      def initialize(sheet, j_row = nil, index = 1)
        self.sheet = sheet
        self.j_row = j_row
        self.index = index
        self.cells = {}
        return unless j_row
        (0..j_row.getLastCellNum.to_i).each do |col|
          j_cell = j_row.getCell(col)
          cells.store col + 1, Worksheet::Cell.new(self, j_cell, col + 1) if j_cell
        end
      end

      def cell(col)
        return cells[col] if cells[col]
        j_cell = j_row.getCell(col)
        cells.store col, Cell.new(self, j_cell, col) if j_cell
      end

      def duplicate_below
        new_row = sheet.create_row(index + 1)
        new_row.copy_row_from(self)
        new_row
      end

      def self.cell_policy
        @cell_policy ||= Rjb.import('org.apache.poi.ss.usermodel.CellCopyPolicy')
      end

      def copy_row_from(row)
        j_row.copyRowFrom(row.j_row, self.class.cell_policy.new)
        (0..j_row.getLastCellNum.to_i).each do |col|
          j_cell = j_row.getCell(col)
          cells.store col + 1, Worksheet::Cell.new(self, j_cell, col + 1) if j_cell
        end
      end

      def inspect
        "row: #{index}, columns: #{cells.values.size}, cell_values: #{cells.values.map(&:val)}"
      end

      def []=(col, value)
        cell(col).val = value
      end

      def [](col)
        return unless cell(col)
        cell(col)
      end

      def cell_values
        cells.values.map(&:val)
      end
    end
  end
end
