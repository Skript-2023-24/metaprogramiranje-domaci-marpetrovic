require "google_drive"

class GoogleSheetTable
  include Enumerable

  attr_reader :headers, :end_col, :start_col

  def initialize(worksheet)
    @worksheet = worksheet
    @header_row_index = find_header_row_index
    @start_col, @end_col = find_table_range
    raise "Header row not found" if @header_row_index.nil?
    raise "Table range not found" if @start_col.nil? || @end_col.nil?
    load_headers
  end

  def row(row_number)
    actual_row_index = @header_row_index + row_number
    raise "Row number is out of range" if actual_row_index > @worksheet.num_rows
  
    @start_col.upto(@end_col).each_with_object([]) do |col, row_values|
      cell_value = @worksheet[actual_row_index, col].to_s
      next if cell_value.empty? || cell_value.match?(/^(total|subtotal)$/i)
  
      row_values << cell_value
    end
  end
  
  def [](column_name)
    col_index = @headers.index(column_name.downcase)
    raise "Column not found: #{column_name}" if col_index.nil?
    
    actual_col_index = col_index + @start_col
    Column.new(@worksheet, @header_row_index, actual_col_index)
  end
  
  def each
    (@header_row_index + 1..@worksheet.num_rows).each do |row_index|
      next if total_or_subtotal?(row_index)
      
      (@start_col..@end_col).each do |col|
        cell_value = @worksheet[row_index, col].to_s
        next if cell_value.empty?
        
        yield cell_value
      end
    end
  end  

  def rows
    (@header_row_index + 1..@worksheet.num_rows).map do |row_index|
      row(row_index - @header_row_index)
    end
  end

  def total_or_subtotal?(row_index)
    (@start_col..@end_col).any? do |col_index|
      cell = @worksheet[row_index, col_index].to_s
      cell.match?(/^(total|subtotal)$/i)
    end
  end

  private

  def load_headers
    @headers = (@start_col..@end_col).map { |col| @worksheet[@header_row_index, col].downcase }
    define_column_methods
  end

  def define_column_methods
    @headers.each_with_index do |header, index|
      define_singleton_method(header.gsub(/\s+/, '')) do |*args|
        column = Column.new(@worksheet, @header_row_index, @start_col + index)
        if args.empty?
          column
        else
          column.find_row_by_value(args.first)
        end
      end
    end
  end

  def find_header_row_index
    (1..@worksheet.num_rows).find do |row|
      (1..@worksheet.num_cols).any? { |col| !@worksheet[row, col].to_s.empty? }
    end
  end

  def find_table_range
    start_col, end_col = nil, nil
    (1..@worksheet.num_cols).each do |col|
      if !@worksheet[@header_row_index, col].to_s.empty?
        start_col ||= col
        end_col = col
      end
    end
    [start_col, end_col]
  end

  class Column
    include Enumerable

    attr_accessor :col_index

  def initialize(worksheet, header_row_index, col_index)
    @worksheet = worksheet
    @header_row_index = header_row_index
    @col_index = col_index
  end
  
  def each
    (@header_row_index + 1..@worksheet.num_rows).each do |row_index|
      next if total_or_subtotal?(row_index)
  
      cell = @worksheet[row_index, @col_index]
      cell_value = cell.to_s
      next if cell_value.empty?
      yield cell_value
    end
  end
  
  def total_or_subtotal?(row_index)
    (1..@worksheet.num_cols).any? do |col_index|
      cell = @worksheet[row_index, col_index].to_s
      cell.match?(/^(total|subtotal)$/i)
    end
  end
  
  def sum
    total_sum = 0.0
    each { |cell_value| total_sum += cell_value.to_f }
    total_sum
  end
  
  def avg
    total_sum = 0.0
    count = 0
    each do |cell_value|
      total_sum += cell_value.to_f
      count += 1
    end
    return 0 if count == 0
    total_sum / count
  end

    def []=(row_index, value)
      actual_row_index = @header_row_index + row_index
      raise "Row index is out of range" if actual_row_index > @worksheet.num_rows or actual_row_index < 1
  
      @worksheet[actual_row_index, @col_index] = value
      @worksheet.save
    end

    def [](row_index)
      actual_row_index = @header_row_index + row_index
      raise "Row index is out of range" if actual_row_index > @worksheet.num_rows or actual_row_index < 1
  
      @worksheet[actual_row_index, @col_index].to_s
    end

    def load_values
      (1..@worksheet.num_rows).map do |row_index|
        @worksheet[row_index, @col_index].to_s
      end
    end

    def find_row_by_value(value)
      (1..@worksheet.num_rows - @header_row_index).each do |row|
        if @worksheet[@header_row_index + row, @col_index].to_s.downcase == value.downcase
          return row
        end
      end
      nil
    end
  end
end