# frozen_string_literal: true

require 'zip'

module Xsv
  class Workbook

    attr_reader :sheets, :shared_strings, :xfs, :numFmts, :trim_empty_rows

    START_OF_DATA = "PK\x03\x04".freeze
    SHARED_STRINGS_PATH = "xl/sharedStrings.xml".freeze
    STYLES_PATH = "xl/styles.xml".freeze
    SHEETS_PATH = "xl/worksheets/sheet*.xml".freeze
    SHEET_NUMBER_PATTERN = /\d+/.freeze
    

    # Open the workbook of the given filename, string or buffer
    def self.open(data, **kws)
      if data.is_a?(IO)
        @workbook = self.new(Zip::InputStream.open(data), kws)
      elsif data.start_with?(START_OF_DATA)
        @workbook = self.new(Zip::InputStream.open(data), kws)
      else
        @workbook = self.new(Zip::InputStream.open(data), kws)
      end
    end

    # Open a workbook from an instance of Zip::File
    #
    # Options:
    #
    #    trim_empty_rows (false) Scan sheet for end of content and don't return trailing rows
    #
    def initialize(zip, trim_empty_rows: false)
      @zip = zip
      @trim_empty_rows = trim_empty_rows

      @sheets = []
      @xfs = []
      @numFmts = Xsv::Helpers::BUILT_IN_NUMBER_FORMATS.dup

      sheets = []
      while (entry = @zip.get_next_entry)
        if entry.name == SHARED_STRINGS_PATH
          @shared_strings = SharedStringsParser.parse(entry.get_input_stream)
        elsif entry.name == STYLES_PATH
          @xfs, @numFmts = StylesHandler.get_styles(entry.get_input_stream, @numFmts)
        elsif entry.name.match?(%r{\Axl/worksheets/sheet.*\.xml\z})
          sheets << entry
        end
      end

      sheets.sort! do |a, b|
        a.name[SHEET_NUMBER_PATTERN].to_i <=> b.name[SHEET_NUMBER_PATTERN].to_i
      end
      sheets.each do |entry|
        @sheets << Xsv::Sheet.new(self, entry.get_input_stream)
      end

      # fetch_shared_strings
      # fetch_styles
      # fetch_sheets
    end

    def inspect
      "#<#{self.class.name}:#{self.object_id}>"
    end

    private

    def fetch_shared_strings
      stream = @zip.glob(SHARED_STRINGS_PATH).first.get_input_stream
      @shared_strings = SharedStringsParser.parse(stream)

      stream.close
    end

    def fetch_styles
      stream = @zip.glob(STYLES_PATH).first.get_input_stream

      @xfs, @numFmts = StylesHandler.get_styles(stream, @numFmts)
    end

    def fetch_sheets
      @zip.glob(SHEETS_PATH).sort do |a, b|
        a.name[SHEET_NUMBER_PATTERN].to_i <=> b.name[SHEET_NUMBER_PATTERN].to_i
      end.each do |entry|
        @sheets << Xsv::Sheet.new(self, entry.get_input_stream)
      end
    end
  end
end
