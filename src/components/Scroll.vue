<template>
    <div class="scroll">
        x: {{posX}}
        <div class="container" @mousemove="mousemove">

        </div>
    </div>
</template>
<script>
export default {
    data() {
        return {
            posX: 0
        }
    },
    methods: {
        mousemove(e) {
            // 获取x 坐标
            this.posX = e.offsetX 
        }
    },   
}
</script>
<style lang="scss" scoped>
.scroll{
    .container{
        width: 960px;
        height: 540px;
        background-color: #cccccc;
    }
}
</style>
