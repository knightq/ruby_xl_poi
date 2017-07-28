# frozen_string_literal: true

module RubyXlPoi
  module Worksheet
    class Sheet
      attr_accessor :rows, :j_sheet, :wb

      def initialize(wb, j_sheet = nil)
        self.wb = wb
        self.j_sheet = j_sheet
        self.rows = {}
        return unless j_sheet
        (0..j_sheet.getPhysicalNumberOfRows.to_i).each do |row_index|
          j_row = j_sheet.getRow(row_index)
          rows.store row_index + 1, Worksheet::Row.new(self, j_row, row_index + 1) if j_row
        end
      end

      def create_row(index)
        Worksheet::Row.new(self, j_sheet.createRow(index), index)
      end

      # Inserts row at row_index, pushes down, copies style from below
      # (row previously at that index) USE OF THIS METHOD will break formulas
      # which reference cells which are being "pushed down".
      def insert_row(row_index = 1)
        ensure_cell_exists(row_index)

        return if row_index < 1
        old_row = rows[row_index - 1]
        old_row&.duplicate_below
      end

      private

      # Ensures that storage space for a cell with +row_index+ and +column_index+
      # exists in +rows+ arrays, growing them up if necessary.
      def ensure_cell_exists(row_index, column_index = 1)
        validate_nonnegative(row_index - 1)
        validate_nonnegative(column_index - 1)

        rows[row_index - 1] || add_row(row_index - 1)
      end

      def validate_nonnegative(row_or_col)
        raise 'Row and Column arguments must be positive' if row_or_col < 1
      end
    end
  end
end
