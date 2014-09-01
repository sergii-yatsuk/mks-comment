import subprocess
from optparse import OptionParser
import re
import os

def RunCmd(path):
   im = subprocess.Popen(path, stdout=subprocess.PIPE)
   return str(im.stdout.read().decode())

class MKSIssue:
   def __init__(self, number):
      self.number = number
      self.im_output = self._imViewissue(number)

   def __str__(self):
      return "MKS Issue: %s\n%s" % (self.number, self.im_output)

   def Number(self):
      return self.number

   def Type(self):
      return self._parseField('Type')

   def Name(self):
      return self._parseField('Name')

   def Description(self):
      return self._parseField('Description')

   def _imViewissue(self, number):
      return RunCmd("im.exe viewissue %s" % number)

   def _parseField(self, field):
      search = re.search(r'^{field}:(?: |)(.*)$'.format(field=field),
                         self.im_output, re.MULTILINE)
      if search is not None:
         return search.groups()[0]
      else:
         raise Exception("Can't find field %s in im.exe output %s"
                         % (field, self.im_output))

class MKSTask(MKSIssue):
   def __init__(self, number):
      super().__init__(number)
      if self.Type() != 'Task':
         raise Exception("MKS issue is not Task: \n%s" % self)

   def ProjectName(self):
      return self._parseField('Project Name')

   def FeatureID(self):
      return self._parseField('Feature ID')


class MKSInspection(MKSIssue):
   def __init__(self, number):
      super().__init__(number)
      if self.Type() != 'Inspection':
         raise Exception("MKS issue is not Inspection: \n%s" % self)

   def Author(self):
      return self._parseField('Author')

   def Moderator(self):
      return self._parseField('Moderator')

   def Ispectors(self):
      return self._parseField('Team Members')


class MKSOutput(MKSIssue):
   def __init__(self, number):
      super().__init__(number)
      if self.Type() != 'Output':
         raise Exception("MKS issue is not Output: \n%s" % self)

   def InspectionCompleted(self):
      return self._parseField('Inspections completed') == 'Yes'

   def GetInspectionNumber(self):
      return self._parseField('Output To Inspection Relationship')

   def GetTaskNumber(self):
      return self._parseField('Output To Task Relationship')

def generateComment(output):
   inspection = MKSInspection(output.GetInspectionNumber())

   task = MKSTask(output.GetTaskNumber())

   return "{output_description}\n\n"\
          "Description :\n"\
          "{output_name}\n"\
          "---Output description---\n"\
          "{output_description}\n"\
          "---Task Description---\n"\
          "{task_name}\n"\
          "MKS Output ID: mks://{output_number}\n"\
          "MKS Feature ID: mks://{feature_id}\n"\
          "MKS Project Name: {project_name}\n"\
          "Reviewed by:  {inspection_number} {moderator}, {inspectors}\n".format(output_name=output.Name(),
                 output_description=output.Description(),
                 task_name=task.Name(),
                 output_number=output.Number(),
                 feature_id=task.FeatureID(),
                 project_name=task.ProjectName(),
                 inspection_number=inspection.Number(),
                 moderator=inspection.Moderator(),
                 inspectors=inspection.Ispectors())

def SearchOutputByHash(hash):
   myOutputs = RunCmd('im.exe issues --query="My outputs" --fields=ID').split()

   for output in myOutputs:
      afterHash = RunCmd("im.exe issues --fields=AfterHash {}".format(output))
      if afterHash.strip() == hash.strip():
         return int(output)

   return None

def CommitBranch(branch):
   commitHash = RunCmd([r"git.exe","rev-parse", branch] )
   output_number = SearchOutputByHash(commitHash)
   output = MKSOutput(output_number)
   RunCmd([r"git.exe","merge", "--squash", branch] )
   RunCmd([r"git.exe","commit", "-m", generateComment(output)] )


def main():
   print(os.getcwd())
   parser = OptionParser(usage="usage: %prog [options]")
   parser.add_option("-o", "--output", type="int",
                     dest="output_number", default=False, metavar="NUM",
                     help="[required] MKS output number")
   parser.add_option("-f", "--force", action="store_true",
                     dest="force", default=False,
                     help="[optional] Ignore not completed inspections")
   parser.add_option("-g", "--hash",
                     dest="commit_hash", default=False, metavar="NUM",
                     help="search output by hash")
   parser.add_option("-b", "--branch",
                     dest="branch", default=False, metavar="NUM",
                     help="commit branch")

   (options, args) = parser.parse_args()


   if options.branch:
      CommitBranch(options.branch)
      return


   output_number = options.output_number

   if options.commit_hash is not None:
      output_number = SearchOutputByHash(options.commit_hash)
      if output_number is None:
         print("Can't find specified has, try generate comment by OutputID")
         return

   if output_number is None:
      parser.print_help()
      return

   output = MKSOutput(output_number)

   if not options.force:
      if not output.InspectionCompleted():
         raise Exception("Inspection is not completed")

   print(generateComment(output))



if __name__ == "__main__":
   main()
