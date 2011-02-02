# This class will allow the automation of word!
#
# For now this module encapsulates word, to generate and manipulate just one word-document
# This could be made more general, so that we could generate more than 1 document at the same time.
# For the moment this is not needed.
#
# So for each document we would want to manipulate/create, we need to create a new WordAutomator instance,
# which would open it's own Word-application object. For the moment that suffices.

require 'win32ole'

class WordAutomator

  attr_reader :version
  attr_reader :word_instance

  def initialize()
    #Launch Microsoft Word:
    @word_instance = WIN32OLE.new('Word.Application')
    @version = @word_instance.Application.Version
    # 9.0 = Word 2000
    # 10.0 = Word 2003
    # 12.0 = Word 2007
  end

  # create a new document, using the template at the given path if given
  def create_document(template_filename=nil)
    raise_if_not_active
    if template_filename.nil?
      @doc = @word.Documents.Add
    else
      @doc = @word.Documents.Add(template_filename)
    end
  end

  # check if the given bookmark exists
  def bookmark_exists?(bm_name)
    raise_if_no_document
    @doc.Bookmarks.Exists(bm_name)
  end

  # this will set text at the bookmark, if the bookmark is a range itself, it will substituted (removed)
  def set_at_bookmark(bm_name, value)
    raise_if_no_document
    # the easy version :
    bm_range = @doc.Bookmarks(bm_name).Range
    bm_range.Text = value
    return bm_range
  end

  # this will set the contents of the bookmark (and will make sure the range of the bookmark is kept --for referring fields a.o.)
  def set_bookmark(bm_name, value)
    raise_if_no_document
    bm_range = set_at_bookmark(bm_name, value)
    @doc.Bookmarks.Add bm_name, bm_range
  end

  # will update all fields inside word (needed so that ref-fields will point to correct bookmarks)
  def update_fields
    raise_if_no_document
    @doc.Fields.Update
  end

  # saves the current document with the given filename
  # this could throw an exception if saving fails!
  def save(filename, as_pdf = false)
    Rails.logger.debug "Saving generated letter in <#{filename}>" unless defined?(Rails).nil?

    version_nr = @version.to_i

    raise_if_no_document
    raise Exception("Saving as PDF is only supported from Word 2007 upwards!") if as_pdf && version_nr.to_i < 12

    if version_nr.to_i < 10
      @word.ActiveDocument.SaveAs filename
    else
      # I assume Word 2003 and 2007 have the same signature
      #
      # specify the word type explicitly, as defined here: http://msdn.microsoft.com/en-us/library/bb238158.aspx
      # 0 or not specified = default save format, as specified from within the program; make to sure to set correctly
      # 16 = default (doc/Word2003)
      # 12 = word xml (docx?) without macros
      # 17 = PDF! :)
      file_ext = File.extname(filename)
      filename = File.join(File.dirname(filename), File.basename(filename, file_ext ) + ".pdf") if as_pdf && file_ext != ".pdf"
      @doc.SaveAs filename, (as_pdf ? 17 : 0)
    end
  end

  # save current document with a temporary filename to save to
  # the temporary name is based on the given model_name (part of the folder)
  # returns the created filename
  def save_temp(model_name, as_pdf=false)
    raise_if_no_document

    retries=0
    begin
      filename = UniqueTempFile.get_filename((as_pdf ? "pdf" : "doc"), "letter", model_name)
      save filename, as_pdf
      return filename
    rescue
      Rails.logger.warn "Saving generated letter in <#{filename}> failed!! Retry (#{retries})" unless defined?(Rails).nil?
      retries+=1
      if retries < 3
        retry
      else
        raise
      end
    end
  end

  # this will close the current document, and close all other open documents (if any)
  # and quit word
  def close_and_quit
    if word_object_is_defined?
      # first close all open documents without saving
      Rails.logger.debug "word_automator::close_and_quit"
      @doc.Close 0 unless @doc.nil?
      1.upto(@word.Documents.Count) do |count|
        @word.Documents.Item(1).Close 0
      end
      Rails.logger.debug "word_automator::quit word"
      @word.Quit
      # clean up, protect against unexpected use
      @word=nil
      @doc=nil
    end
  end

  private

  def raise_if_not_active
    raise Exception("The WORD-object is no longer active! Hint: have you called close_and_quit before?") unless word_object_is_defined?
  end

  def raise_if_no_document
    raise_if_not_active
    raise Exception("There is no known active document! Hint: have you called create_document first?") if @doc.nil?
  end

  def word_object_is_defined?
    !@word.nil?
  end


end