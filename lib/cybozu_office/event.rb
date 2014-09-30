module CybozuOffice
  class Event
    def initialize
      @eid = ''
      @title = ''
      @year = ''
      @month =''
      @day = ''
      @hour_b = ''
      @minute_b = ''
      @hour_e = ''
      @minute_e = ''
      @updated_at = ''
    end

    def eid
      return @eid
    end
    def set_eid(int)
      @eid = int.to_i
      return @self
    end

    def title
      return @title
    end
    def set_title(str)
      @title = str.to_s
      return @self
    end

    def year
      return @year
    end
    def set_year(str)
      @year = str.to_s
      return @self
    end

    def month
      return @month
    end
    def set_month(str)
      @month = str.to_s
      return @self
    end

    def day
      return @day
    end
    def set_day(str)
      @day = str.to_s
      return @self
    end

    def hour_b
      return @hour_b
    end
    def set_hour_b(str)
      @hour_b = str.to_s
      return @self
    end

    def minute_b
      return @minute_b
    end
    def set_minute_b(str)
      @minute_b = str.to_s
      return @self
    end

    def hour_e
      return @hour_e
    end
    def set_hour_e(str)
      @hour_e = str.to_s
      return @self
    end

    def minute_e
      return @minute_e
    end
    def set_minute_e(str)
      @minute_e = str.to_s
      return @self
    end

    def get_updated_at
      return @updated_at
    end
    def set_updated_at(updated_at)
      @updated_at = updated_at
      return @self
    end
  end
end

