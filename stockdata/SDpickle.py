import pickle

class SDPickleMixin():
      def pickle_dump(self, filename):
          pickle.dump(self.data, open(filename, "wb"))

      def pickle_read(self, filename):
          pload = True
          try:
             self.data = pickle.load(open(filename, "rb"))
          except (OSError, IOError) as e:
             pload = False
          return pload
