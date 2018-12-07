#/bin/sh
# ./protoc -I=./include/google/protobuf --cpp_out=./game ./setting/skill.proto
./protoc -I=./setting --python_out=./game  ./setting/skill.proto
