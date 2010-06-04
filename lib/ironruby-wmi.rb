require 'System';
require 'System.Management, Version=2.0.0.0, 
     Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'

class String
  def underscore
    scan(/[A-Z][a-z]+/).to_a.map(&:downcase).join("_")
  end
end

module WMI
  class WMIError <StandardError
  end

  class InvalidQuery <WMIError
  end

  class Base
    def initialize obj
      @obj = obj
    end

    def wmi_delegate_obj
      @obj
    end

    def self.wmi_delegate_obj
      wmi_class
    end

    def self.wmi_class
      @wmi_class ||= System::Management::ManagementClass.new(class_name)
    end

    def self.class_name
      @class_name ||= self.name.split("::").last
    end

    def self.get_instances
      wmi_class.get_instances.map do |obj|
        self.new(obj)
      end
    end

    def self.get_instances_async &block
      searcher = System::Management::ManagementObjectSearcher.new(
        System::Management::SelectQuery.new(class_name));
      results = System::Management::ManagementOperationObserver.new

      m = Module.new
      m.module_eval do
        define_method :call, block
      end

      klass = self
      results.ObjectReady do |observer, event_args|
        instance = klass.new(event_args.NewObject)
        instance.extend m
        instance.call instance
      end
      completed = false
      results.Completed do |observer, event_args|
        completed = true
      end
      at_exit do
        sleep 1 until completed
      end
      searcher.Get(results)
    end

    def check_return_value return_value
      self.class.check_return_value return_value
    end

    def self.check_return_value return_value
      case return_value
      when 0
        # do nothing
      when 2
        raise WMI::WMIError.new("Access Denied");
      when 3
        raise WMI::WMIError.new("Insufficient Privilege");
      when 8
        raise WMI::WMIError.new("Unknown failure");
      when 9
        raise WMI::WMIError.new("Path Not Found");
      when 21
        raise WMI::WMIError.new("Invalid Parameter");
      else
        raise WMI::WMIError.new("ReturnValue == " + return_value)
      end
    end

    def invoke_method_1 method_name, *args
      return_value = @obj.InvokeMethod method_name, args
      self.class.check_return_value return_value
    end

    def method_missing name, *args
      @obj.__send__ name, *args
    end
  end


  def self.const_missing name
    klass = Class.new(self::Base)
    self.const_set(name, klass)
    klass.class_eval do
      wmi_class.Properties.each do |prop|
        prop_name = prop.Name
        if prop.IsArray
          define_method prop_name do
            @obj.Properties[prop_name].Value.to_a
          end
        else
          define_method prop_name do
            @obj.Properties[prop_name].Value
          end
        end
        if prop_name != prop_name.underscore then
          alias_method prop_name.underscore.to_sym, prop_name.to_sym
        end
      end
    end

    klass.wmi_class.Methods.each do |m|
      method_name = m.Name.to_s

      param_count = out_count = m.OutParameters.Properties.Count - 1
      if m.InParameters
        param_count += m.InParameters.Properties.Count
      end
      mm = %(
        def #{method_name} *args
          (#{param_count} - args.size).times do
            args.push nil
          end
          array = args.ToArray
          return_value = wmi_delegate_obj.InvokeMethod "#{method_name}", array
          check_return_value return_value
          if #{out_count} == 0 then
            return nil
          else
            return *array.to_a[-#{out_count} .. -1]
          end
        end
      )
      klass.instance_eval mm
      klass.class_eval mm
      if method_name != method_name.underscore then
        mmm = %{
          class <<self
            alias :#{method_name.underscore} :#{method_name}
          end
        }
        klass.instance_eval mmm
        klass.class_eval    mmm
      end
    end
    code = %(
      def method_missing method_name, *args
        if method_name == :InvokeMethod then
          wmi_method = args.shift.ToString
          self.__send__ wmi_method, *args
        else
          result = wmi_delegate_obj.__send__ method_name, *args
          if result.is_a? (System::Management::ManagementObjectCollection) then
              return result.map do |obj|
                  klass.new obj
              end
          end
          return result
        end
      end
    )
    klass.instance_eval code
    klass.class_eval code
    klass
  end
end

if $0 == __FILE__

end

