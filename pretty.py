import pickle 
import pprint
pp = pprint.PrettyPrinter(indent=4)
dicc = pickle.load( open( "dicc.p", "rb" ) )
pp.pprint(dicc)


print ("PRINTING UNIQUE")

unique = pickle.load( open( "unique.p", "rb" ) )
pp.pprint(unique)