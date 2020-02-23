# frozen_string_literal: true
module Xsv
  class SheetRowsHandler < Ox::Sax
    include Xsv::Helpers

    CELL_TYPE_SHARED_STRING = "s".freeze
    CELL_TYPE_STRING = "str".freeze
    CELL_TYPE_EMPTY = "e".freeze

    def format_cell
      case @current_cell[:t]
      when CELL_TYPE_SHARED_STRING
        @workbook.shared_strings[@current_value.to_i]
      when CELL_TYPE_STRING
        @current_value
      when CELL_TYPE_EMPTY # N/A
        nil
      when nil
        if @current_value.empty?
          nil
        elsif (string = @current_cell[:s])
          style = @workbook.xfs[string.to_i]
          numFmt = @workbook.numFmts[style[:numFmtId].to_i]

          parse_number_format(@current_value, numFmt)
        else
          parse_number(@current_value)
        end
      else
        raise Xsv::Error, "Encountered unknown column type #{@current_cell[:t]}"
      end
    end

    # Ox::Sax implementation below

    def initialize(mode, empty_row, workbook, row_skip, last_row, &block)
      @block = block

      # :sheetData
      # :row
      # :c
      # :v
      @state = nil

      @mode = mode
      @empty_row = empty_row
      @workbook = workbook
      @row_skip = row_skip
      @row_index = 0 - @row_skip
      @current_row = {}
      @current_row_attrs = {}
      @current_cell = {}
      @current_value = nil
      @last_row = last_row

      if @mode == :hash
        @headers = @empty_row.keys
      end
    end

    def start_element(name)
      case name
      when :c
        @state = name
        @current_cell = {}
        @current_value = String.new
      when :v
        @state = name
      when :row
        @state = name
        @current_row = @empty_row.dup
        @current_row_attrs = {}
      else
        @state = nil
      end
    end

    def text(value)
      if @state == :v
        @current_value << value
      end
    end

    def attr(name, value)
      case @state
      when :c
        @current_cell[name] = value
      when :row
        @current_row_attrs[name] = value
      end
    end

    def end_element(name)
      case name
      when :c
        col_index = column_index(@current_cell[:r])

        case @mode
        when :array
          @current_row[col_index] = format_cell
        when :hash
          @current_row[@headers[col_index]] = format_cell
        end
      when :row
        if @row_index < 0
          @row_index += 1
          return
        end

        @row_index += 1

        # Skip first row if we're in hash mode
        return if @row_index == 1 && @mode == :hash

        # Pad empty rows
        while @row_index < @current_row_attrs[:r].to_i - @row_skip
          @block.call(@empty_row)
          @row_index += 1
        end

        # Do not return empty trailing rows
        @block.call(@current_row) unless @row_index > @last_row - @row_skip
      end
    end
  end
end
